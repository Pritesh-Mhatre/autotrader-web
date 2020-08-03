Attribute VB_Name = "AutoTraderWeb"
Option Explicit

Dim ORDER_NUM As Integer
Dim START_TIME As Long

Const CONTACT_SUPPORT As String = "Please take a screenshot of this message and mail to help@stocksdeveloper.in"

Const COMMANDS_FILE As String = "commands.csv"

Const INPUT_DIR As String = "input"
Const OUTPUT_DIR As String = "output"

Const CANCEL_ORDER_CMD As String = "CANCEL_ORDER"
Const MODIFY_ORDER_CMD As String = "MODIFY_ORDER"
Const CANCEL_CHILD_ORDER_CMD As String = "CANCEL_CHILD_ORDER"

Public Const EPOCH As Date = #1/1/1970#
Public Const BLANK As String = ""

Public Const VARIETY_REGULAR As String = "REGULAR"
Public Const VARIETY_BO As String = "BO"
Public Const VARIETY_CO As String = "CO"

Public Const VALIDITY_DAY As String = "DAY"
Public Const VALIDITY_IOC As String = "IOC"
Public Const VALIDITY_DEFAULT As String = VALIDITY_DAY

Public Const PRODUCT_INTRADAY As String = "INTRADAY"
Public Const PRODUCT_DELIVERY As String = "DELIVERY"
Public Const PRODUCT_NORMAL As String = "NORMAL"

Private Sub Sleep(seconds As Integer)

    Dim newHour As Integer
    Dim newMinute As Integer
    Dim newSecond As Integer
    Dim waitTime As Date

    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + seconds
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    
    Application.Wait waitTime

End Sub

Private Function GetIPCDirectory() As String
    
    GetIPCDirectory = Environ("USERPROFILE") & Application.PathSeparator & "autotrader"

End Function

Private Function GetOutputDirectory() As String
    
    GetOutputDirectory = GetIPCDirectory & Application.PathSeparator _
        & OUTPUT_DIR

End Function

Private Function GetCommandsFilePath() As String
    
    GetCommandsFilePath = GetIPCDirectory & Application.PathSeparator _
        & INPUT_DIR & Application.PathSeparator & COMMANDS_FILE

End Function

Public Function GetPortfolioOrdersFile(pseudoAccount As String) As String

	GetPortfolioOrdersFile = GetOutputDirectory & Application.PathSeparator _
		&  pseudoAccount & "-orders.csv"

End Function

Public Function GetPortfolioPositionsFile(pseudoAccount As String) As String

	GetPortfolioPositionsFile = GetOutputDirectory & Application.PathSeparator _
		&  pseudoAccount & "-positions.csv"

End Function

Public Function GetPortfolioMarginsFile(pseudoAccount As String) As String

	GetPortfolioMarginsFile = GetOutputDirectory & Application.PathSeparator _
		&  pseudoAccount & "-margins.csv"

End Function

Public Function FileReadCsvColumnByRowId(filePath As String, _
	rowId As String, rowIdColumnIndex As Integer, columnIndex As Integer) As String
    
    On Error GoTo Done
    
    Dim temp As String
    Dim cols() As String

    FileReadCsvColumnByRowId = ""
    
    Open filePath For Input As #1
    
    Do Until EOF(1)
        Line Input #1, temp
        cols = Split(temp, ",")
        
        If cols(rowIdColumnIndex - 1) = rowId Then
            FileReadCsvColumnByRowId = cols(columnIndex - 1)
            Exit Do
        End If
        
    Loop
    
    Close #1

Done:
    Exit Function
    
End Function

Private Function NextOrderNumber() As String

    If ORDER_NUM = 0 Then
        START_TIME = Abs(CLng((Now() - EPOCH) * 86400 - 2 ^ 31))
    End If
    
    ORDER_NUM = ORDER_NUM + 1
    NextOrderNumber = CStr(START_TIME + ORDER_NUM)
    
End Function

Private Function ValidateFile(FilePath As String, Message As String) As Boolean

    With (CreateObject("Scripting.FileSystemObject"))
        If Not .FileExists(FilePath) Then
            MsgBox Message, vbCritical, "Error"
            ValidateFile = False
        Else
            ValidateFile = True
        End If
    End With

End Function

Private Sub WriteCommand(Command As String)

    Dim CommandsFilePath As String
    
    CommandsFilePath = GetCommandsFilePath()
    If ValidateFile(CommandsFilePath, "AutoTrader client is not monitoring commands file.") = False Then
        Exit Sub
    End If

    Open CommandsFilePath For Append As #1
        Print #1, Command
    Close #1

End Sub

Private Function PlaceOrderInternal(Order As Order) As String
    
    On Error GoTo Error_Handler

    ' Assign a unique order id
    Order.PublisherId = NextOrderNumber()
    
    ' Write PlaceOrder Command to File
    WriteCommand (Order.ToPlaceCommand)

    PlaceOrderInternal = Order.PublisherId

Error_Handler:
    If Err.Number <> 0 Then
        
        Dim Message As String
        Message = "The Error Happened on Line : " & Erl & vbNewLine & _
                        "Error Message : " & Err.Description & vbNewLine & _
                        "Error Number : " & Err.Number & vbNewLine & vbNewLine & _
                        CONTACT_SUPPORT
                
        MsgBox Message, vbOKOnly, "Error"
        Resume Next

    End If
    
End Function

Public Function PlaceOrderAdvanced(Variety As String, _
    PseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    ProductType As String, _
    Quantity As Integer, _
    Price As Double, _
    TriggerPrice As Double, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double, _
    DisclosedQuantity As Integer, _
    Validity As String, _
    Amo As Boolean, _
    StrategyId As Integer, _
    Comments As String) As String

    Dim o As New Order
    
    Dim TmpVariety As New Variety
    Dim TmpTradeType As New TradeType
    Dim TmpOrderType As New OrderType
    Dim TmpProductType As New ProductType
    Dim TmpValidity As New Validity
    
    TmpVariety.FromString (Variety)
    o.Variety = TmpVariety
    
    TmpTradeType.FromString (TradeType)
    o.TradeType = TmpTradeType
    
    TmpOrderType.FromString (OrderType)
    o.OrderType = TmpOrderType
    
    TmpProductType.FromString (ProductType)
    o.ProductType = TmpProductType
    
    TmpValidity.FromString (Validity)
    o.Validity = TmpValidity
    
    o.PseudoAccount = PseudoAccount
    o.Exchange = Exchange
    o.Symbol = Symbol
    o.Quantity = Quantity
    o.Price = Price
    o.TriggerPrice = TriggerPrice
    o.Amo = Amo
    o.DisclosedQuantity = DisclosedQuantity
    o.Target = Target
    o.Stoploss = Stoploss
    o.TrailingStoploss = TrailingStoploss
    o.Comments = Comments
    o.StrategyId = StrategyId

    PlaceOrderAdvanced = PlaceOrderInternal(o)

End Function
        
Public Function PlaceOrder( _
    PseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    ProductType As String, _
    Quantity As Integer, _
    Price As Double, _
    TriggerPrice As Double) As String
        
    Dim Variety As String: Variety = VARIETY_REGULAR
    Dim Target As Double: Target = 0
    Dim Stoploss As Double: Stoploss = 0
    Dim TrailingStoploss As Double: TrailingStoploss = 0
    Dim DisclosedQuantity As Integer: DisclosedQuantity = 0
    Dim Validity As String: Validity = VALIDITY_DEFAULT
    Dim Amo As Boolean: Amo = False
    Dim StrategyId As Integer: StrategyId = -1
    Dim Comments As String: Comments = ""
        
    PlaceOrder = PlaceOrderAdvanced(Variety, PseudoAccount, _
            Exchange, Symbol, TradeType, OrderType, ProductType, _
            Quantity, Price, TriggerPrice, Target, Stoploss, _
            TrailingStoploss, DisclosedQuantity, Validity, _
            Amo, StrategyId, Comments)

End Function
        
Public Function PlaceBracketOrder( _
    PseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    Quantity As Integer, _
    Price As Double, _
    TriggerPrice As Double, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double) As String

    Dim Variety As String: Variety = VARIETY_BO
    Dim DisclosedQuantity As Integer: DisclosedQuantity = 0
    Dim Validity As String: Validity = VALIDITY_DEFAULT
    Dim Amo As Boolean: Amo = False
    Dim StrategyId As Integer: StrategyId = -1
    Dim Comments As String: Comments = ""
    Dim ProductType As String: ProductType = PRODUCT_INTRADAY

    PlaceBracketOrder = PlaceOrderAdvanced(Variety, PseudoAccount, _
            Exchange, Symbol, TradeType, OrderType, ProductType, _
            Quantity, Price, TriggerPrice, Target, Stoploss, _
            TrailingStoploss, DisclosedQuantity, Validity, _
            Amo, StrategyId, Comments)

End Function

Public Function PlaceCoverOrder( _
    PseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    Quantity As Integer, _
    Price As Double, _
    TriggerPrice As Double) As String

    Dim Variety As String: Variety = VARIETY_CO
    Dim DisclosedQuantity As Integer: DisclosedQuantity = 0
    Dim Validity As String: Validity = VALIDITY_DEFAULT
    Dim Amo As Boolean: Amo = False
    Dim StrategyId As Integer: StrategyId = -1
    Dim Comments As String: Comments = ""
    Dim ProductType As String: ProductType = PRODUCT_INTRADAY
    Dim Target As Double: Target = 0
    Dim Stoploss As Double: Stoploss = 0
    Dim TrailingStoploss As Double: TrailingStoploss = 0

    PlaceCoverOrder = PlaceOrderAdvanced(Variety, PseudoAccount, _
            Exchange, Symbol, TradeType, OrderType, ProductType, _
            Quantity, Price, TriggerPrice, Target, Stoploss, _
            TrailingStoploss, DisclosedQuantity, Validity, _
            Amo, StrategyId, Comments)

End Function

Public Function CancelOrder(PseudoAccount As String, _
        OrderId As String) As Boolean

    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 2) As String

    cols(0) = CANCEL_ORDER_CMD
    cols(1) = PseudoAccount
    cols(2) = OrderId

    csv = Join(cols, ",")
        
    WriteCommand (csv)
        
    CancelOrder = True

Error_Handler:
    If Err.Number <> 0 Then
        
        Dim Message As String
        Message = "The Error Happened on Line : " & Erl & vbNewLine & _
                        "Error Message : " & Err.Description & vbNewLine & _
                        "Error Number : " & Err.Number & vbNewLine & vbNewLine & _
                        CONTACT_SUPPORT
                
        MsgBox Message, vbOKOnly, "Error"
        Resume Next

    End If
        
End Function

Public Function CancelOrderChildren(PseudoAccount As String, _
        OrderId As String) As Boolean

    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 2) As String

    cols(0) = CANCEL_CHILD_ORDER_CMD
    cols(1) = PseudoAccount
    cols(2) = OrderId

    csv = Join(cols, ",")
        
    WriteCommand (csv)
        
    CancelOrderChildren = True

Error_Handler:
    If Err.Number <> 0 Then
        
        Dim Message As String
        Message = "The Error Happened on Line : " & Erl & vbNewLine & _
                        "Error Message : " & Err.Description & vbNewLine & _
                        "Error Number : " & Err.Number & vbNewLine & vbNewLine & _
                        CONTACT_SUPPORT
                
        MsgBox Message, vbOKOnly, "Error"
        Resume Next

    End If
        
End Function

Public Function ModifyOrder(PseudoAccount As String, _
    OrderId As String, _
    OrderType As String, _
    Quantity As Integer, _
    Price As Double, _
    TriggerPrice As Double) As Boolean
        
    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 6) As String

    cols(0) = MODIFY_ORDER_CMD
    cols(1) = PseudoAccount
    cols(2) = OrderId
    cols(3) = OrderType
    cols(4) = Quantity
    cols(5) = Price
    cols(6) = TriggerPrice

    csv = Join(cols, ",")
        
    WriteCommand (csv)
        
    ModifyOrder = True

Error_Handler:
    If Err.Number <> 0 Then
        
        Dim Message As String
        Message = "The Error Happened on Line : " & Erl & vbNewLine & _
                        "Error Message : " & Err.Description & vbNewLine & _
                        "Error Number : " & Err.Number & vbNewLine & vbNewLine & _
                        CONTACT_SUPPORT
                
        MsgBox Message, vbOKOnly, "Error"
        Resume Next

    End If
        
End Function

Public Function ModifyOrderPrice(PseudoAccount As String, _
    OrderId As String, _
    Price As Double) As Boolean
        
    ModifyOrderPrice = ModifyOrder(PseudoAccount, OrderId, "", 0, Price, 0)

End Function

Public Function ModifyOrderQuantity(PseudoAccount As String, _
    OrderId As String, _
    Quantity As Integer) As Boolean

    ModifyOrderQuantity = ModifyOrder(PseudoAccount, OrderId, "", Quantity, 0, 0)

End Function

Public Function isAutoTraderClientMonitoring() As Boolean
    
    Dim CommandsFilePath As String
    
    CommandsFilePath = GetCommandsFilePath()
    isAutoTraderClientMonitoring = ValidateFile(CommandsFilePath, _
        "AutoTrader client is not monitoring commands file.")

End Function

' *****************************************************************************/
' ************************ ORDER DETAIL FUNCTIONS - END ***********************/
' *****************************************************************************/

' Reads orders file and returns a column value for the given order id.
Public Function ReadOrderColumn(pseudoAccount As String, _
	orderId As String, columnIndex As Integer) As String
    Dim filePath As String
	filePath = GetPortfolioOrdersFile(pseudoAccount)
    ReadOrderColumn = FileReadCsvColumnByRowId( filePath, orderId, 3, columnIndex )
End Function

' Retrieve order's trading account.
Public Function GetOrderTradingAccount(pseudoAccount As String, _
	orderId As String) As String
	GetOrderTradingAccount = ReadOrderColumn(pseudoAccount, orderId, 2)
End Function

' Retrieve order's trading platform id.
Public Function GetOrderId(pseudoAccount As String, _
	orderId As String) As String
	GetOrderId = ReadOrderColumn(pseudoAccount, orderId, 4)
End Function

' Retrieve order's exchange id.
Public Function GetOrderExchangeId(pseudoAccount As String, _
	orderId As String) As String
	GetOrderExchangeId = ReadOrderColumn(pseudoAccount, orderId, 5)
End Function

' Retrieve order's variety (REGULAR, BO, CO).
Public Function GetOrderVariety(pseudoAccount As String, _
	orderId As String) As String
	GetOrderVariety = ReadOrderColumn(pseudoAccount, orderId, 6)
End Function

' Retrieve order's (platform independent) exchange.
Public Function GetOrderIndependentExchange(pseudoAccount As String, _
	orderId As String) As String
	GetOrderIndependentExchange = ReadOrderColumn(pseudoAccount, orderId, 7)
End Function

' Retrieve order's (platform independent) symbol.
Public Function GetOrderIndependentSymbol(pseudoAccount As String, _
	orderId As String) As String
	GetOrderIndependentSymbol = ReadOrderColumn(pseudoAccount, orderId, 8)
End Function

' Retrieve order's trade type (BUY, SELL).
Public Function GetOrderTradeType(pseudoAccount As String, _
	orderId As String) As String
	GetOrderTradeType = ReadOrderColumn(pseudoAccount, orderId, 9)
End Function

' Retrieve order's order type (LIMIT, MARKET, STOP_LOSS, SL_MARKET).
Public Function GetOrderOrderType(pseudoAccount As String, _
	orderId As String) As String
	GetOrderOrderType = ReadOrderColumn(pseudoAccount, orderId, 10)
End Function

' Retrieve order's product type (INTRADAY, DELIVERY, NORMAL).
Public Function GetOrderProductType(pseudoAccount As String, _
	orderId As String) As String
	GetOrderProductType = ReadOrderColumn(pseudoAccount, orderId, 11)
End Function

' Retrieve order's quantity.
Public Function GetOrderQuantity(pseudoAccount As String, _
	orderId As String) As Long
	GetOrderQuantity = CLng(ReadOrderColumn(pseudoAccount, orderId, 12))
End Function

' Retrieve order's price.
Public Function GetOrderPrice(pseudoAccount As String, _
	orderId As String) As Double
	GetOrderPrice = CDbl(ReadOrderColumn(pseudoAccount, orderId, 13))
End Function

' Retrieve order's trigger price.
Public Function GetOrderTriggerPrice(pseudoAccount As String, _
	orderId As String) As Double
	GetOrderTriggerPrice = CDbl(ReadOrderColumn(pseudoAccount, orderId, 14))
End Function

' Retrieve order's filled quantity.
Public Function GetOrderFilledQuantity(pseudoAccount As String, _
	orderId As String) As Long
	GetOrderFilledQuantity = CLng(ReadOrderColumn(pseudoAccount, orderId, 15))
End Function

' Retrieve order's pending quantity.
Public Function GetOrderPendingQuantity(pseudoAccount As String, _
	orderId As String) As Long
	GetOrderPendingQuantity = CLng(ReadOrderColumn(pseudoAccount, orderId, 16))
End Function

' Retrieve order's (platform independent) status.
' (OPEN, COMPLETE, CANCELLED, REJECTED, TRIGGER_PENDING, UNKNOWN)
Public Function GetOrderStatus(pseudoAccount As String, _
	orderId As String) As String
	GetOrderStatus = ReadOrderColumn(pseudoAccount, orderId, 17)
End Function

' Retrieve order's status message or rejection reason.
Public Function GetOrderStatusMessage(pseudoAccount As String, _
	orderId As String) As String
	GetOrderStatusMessage = ReadOrderColumn(pseudoAccount, orderId, 18)
End Function

' Retrieve order's validity (DAY, IOC).
Public Function GetOrderValidity(pseudoAccount As String, _
	orderId As String) As String
	GetOrderValidity = ReadOrderColumn(pseudoAccount, orderId, 19)
End Function

' Retrieve order's average price at which it got traded.
Public Function GetOrderAveragePrice(pseudoAccount As String, _
	orderId As String) As Double
	GetOrderAveragePrice = CDbl(ReadOrderColumn(pseudoAccount, orderId, 20))
End Function

' Retrieve order's parent order id. The id of parent bracket or cover order.
Public Function GetOrderParentOrderId(pseudoAccount As String, _
	orderId As String) As String
	GetOrderParentOrderId = ReadOrderColumn(pseudoAccount, orderId, 21)
End Function

' Retrieve order's disclosed quantity.
Public Function GetOrderDisclosedQuantity(pseudoAccount As String, _
	orderId As String) As Long
	GetOrderDisclosedQuantity = CLng(ReadOrderColumn(pseudoAccount, orderId, 22))
End Function

' Retrieve order's exchange time as a string (YYYY-MM-DD HH:MM:SS.MILLIS).
Public Function GetOrderExchangeTime(pseudoAccount As String, _
	orderId As String) As String
	GetOrderExchangeTime = ReadOrderColumn(pseudoAccount, orderId, 23)
End Function

' Retrieve order's platform time as a string (YYYY-MM-DD HH:MM:SS.MILLIS).
Public Function GetOrderPlatformTime(pseudoAccount As String, _
	orderId As String) As String
	GetOrderPlatformTime = ReadOrderColumn(pseudoAccount, orderId, 24)
End Function

' Retrieve order's AMO (after market order) flag. (true/false)
Public Function GetOrderAmo(pseudoAccount As String, _
	orderId As String) As Boolean
	Dim flag As String
	flag = ReadOrderColumn(pseudoAccount, orderId, 25)
	GetOrderAmo = (LCase(flag) = "true")
End Function

' Retrieve order's comments.
Public Function GetOrderComments(pseudoAccount As String, _
	orderId As String) As String
	GetOrderComments = ReadOrderColumn(pseudoAccount, orderId, 26)
End Function

' Retrieve order's raw (platform specific) status.
Public Function GetOrderRawStatus(pseudoAccount As String, _
	orderId As String) As String
	GetOrderRawStatus = ReadOrderColumn(pseudoAccount, orderId, 27)
End Function

' Retrieve order's (platform specific) exchange.
Public Function GetOrderExchange(pseudoAccount As String, _
	orderId As String) As String
	GetOrderExchange = ReadOrderColumn(pseudoAccount, orderId, 28)
End Function

' Retrieve order's (platform specific) symbol.
Public Function GetOrderSymbol(pseudoAccount As String, _
	orderId As String) As String
	GetOrderSymbol = ReadOrderColumn(pseudoAccount, orderId, 29)
End Function

' Retrieve order's date (DD-MM-YYYY).
Public Function GetOrderDay(pseudoAccount As String, _
	orderId As String) As String
	GetOrderDay = ReadOrderColumn(pseudoAccount, orderId, 30)
End Function

' Retrieve order's trading platform.
Public Function GetOrderPlatform(pseudoAccount As String, _
	orderId As String) As String
	GetOrderPlatform = ReadOrderColumn(pseudoAccount, orderId, 31)
End Function

' Retrieve order's client id (as received from trading platform).
Public Function GetOrderClientId(pseudoAccount As String, _
	orderId As String) As String
	GetOrderClientId = ReadOrderColumn(pseudoAccount, orderId, 32)
End Function

' Retrieve order's stock broker.
Public Function GetOrderStockBroker(pseudoAccount As String, _
	orderId As String) As String
	GetOrderStockBroker = ReadOrderColumn(pseudoAccount, orderId, 33)
End Function
