Attribute VB_Name = "AutoTraderWeb"
Option Explicit

Dim ORDER_NUM As Integer
Dim START_TIME As Long

Const CONTACT_SUPPORT As String = "Please take a screenshot of this message and mail to help@stocksdeveloper.in"

Const COMMANDS_FILE As String = "commands.csv"

Const INPUT_DIR As String = "input"
Const OUTPUT_DIR As String = "output"

Const CANCEL_ORDER_CMD As String = "CANCEL_ORDER"
Const CANCEL_ALL_ORDERS_CMD As String = "CANCEL_ALL_ORDERS"
Const MODIFY_ORDER_CMD As String = "MODIFY_ORDER"
Const CANCEL_CHILD_ORDER_CMD As String = "CANCEL_CHILD_ORDER"
Const SQUARE_OFF_POSITION_CMD As String = "SQUARE_OFF_POSITION"
Const SQUARE_OFF_PORTFOLIO_CMD As String = "SQUARE_OFF_PORTFOLIO"

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
Public Const PRODUCT_MTF As String = "MTF"

Public Const MARGIN_EQUITY As String = "EQUITY"
Public Const MARGIN_COMMODITY As String = "COMMODITY"
Public Const MARGIN_ALL As String = "ALL"

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

Function FileExists(FilePath As String) As Boolean
    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

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
        & pseudoAccount & "-orders.csv"

End Function

Public Function GetPortfolioPositionsFile(pseudoAccount As String) As String

    GetPortfolioPositionsFile = GetOutputDirectory & Application.PathSeparator _
        & pseudoAccount & "-positions.csv"

End Function

Public Function GetPortfolioMarginsFile(pseudoAccount As String) As String

    GetPortfolioMarginsFile = GetOutputDirectory & Application.PathSeparator _
        & pseudoAccount & "-margins.csv"

End Function

Public Function GetPortfolioSummaryFile(pseudoAccount As String) As String

    GetPortfolioSummaryFile = GetOutputDirectory & Application.PathSeparator _
        & pseudoAccount & "-summary.csv"

End Function

Public Function FileReadCsvColumnByRowId(FilePath As String, _
    rowId As String, rowIdColumnIndex As Integer, columnIndex As Integer) As String
    
    On Error GoTo Done
    
    Dim temp As String
    Dim cols() As String

    FileReadCsvColumnByRowId = ""
    
    Open FilePath For Input As #1
    
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

Private Function NextOrderNumber() As String

    If ORDER_NUM = 0 Then
        START_TIME = Abs(CLng((Now() - EPOCH) * 86400 - 2 ^ 31))
    End If
    
    ORDER_NUM = ORDER_NUM + 1
    NextOrderNumber = CStr(START_TIME + ORDER_NUM)
    
End Function

Private Function ValidateFile(FilePath As String, Message As String) As Boolean

    If FileExists(FilePath) Then
        ValidateFile = True
    Else
        MsgBox Message, vbCritical, "Error"
        ValidateFile = False
    End If

End Function

Private Sub WriteCommand(Command As String)

    Dim CommandsFilePath As String
    
    CommandsFilePath = GetCommandsFilePath()
    If ValidateFile(CommandsFilePath, "AutoTrader client is not monitoring commands file. " & CommandsFilePath) = False Then
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
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    ProductType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double, _
    DisclosedQuantity As Long, _
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
    
    o.pseudoAccount = pseudoAccount
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
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    ProductType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double) As String
        
    Dim Variety As String: Variety = VARIETY_REGULAR
    Dim Target As Double: Target = 0
    Dim Stoploss As Double: Stoploss = 0
    Dim TrailingStoploss As Double: TrailingStoploss = 0
    Dim DisclosedQuantity As Long: DisclosedQuantity = 0
    Dim Validity As String: Validity = VALIDITY_DEFAULT
    Dim Amo As Boolean: Amo = False
    Dim StrategyId As Integer: StrategyId = -1
    Dim Comments As String: Comments = ""
        
    PlaceOrder = PlaceOrderAdvanced(Variety, pseudoAccount, _
            Exchange, Symbol, TradeType, OrderType, ProductType, _
            Quantity, Price, TriggerPrice, Target, Stoploss, _
            TrailingStoploss, DisclosedQuantity, Validity, _
            Amo, StrategyId, Comments)

End Function
        
Public Function PlaceBracketOrder( _
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double) As String

    Dim Variety As String: Variety = VARIETY_BO
    Dim DisclosedQuantity As Long: DisclosedQuantity = 0
    Dim Validity As String: Validity = VALIDITY_DEFAULT
    Dim Amo As Boolean: Amo = False
    Dim StrategyId As Integer: StrategyId = -1
    Dim Comments As String: Comments = ""
    Dim ProductType As String: ProductType = PRODUCT_INTRADAY

    PlaceBracketOrder = PlaceOrderAdvanced(Variety, pseudoAccount, _
            Exchange, Symbol, TradeType, OrderType, ProductType, _
            Quantity, Price, TriggerPrice, Target, Stoploss, _
            TrailingStoploss, DisclosedQuantity, Validity, _
            Amo, StrategyId, Comments)

End Function

Public Function PlaceCoverOrder( _
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double) As String

    Dim Variety As String: Variety = VARIETY_CO
    Dim DisclosedQuantity As Long: DisclosedQuantity = 0
    Dim Validity As String: Validity = VALIDITY_DEFAULT
    Dim Amo As Boolean: Amo = False
    Dim StrategyId As Integer: StrategyId = -1
    Dim Comments As String: Comments = ""
    Dim ProductType As String: ProductType = PRODUCT_INTRADAY
    Dim Target As Double: Target = 0
    Dim Stoploss As Double: Stoploss = 0
    Dim TrailingStoploss As Double: TrailingStoploss = 0

    PlaceCoverOrder = PlaceOrderAdvanced(Variety, pseudoAccount, _
            Exchange, Symbol, TradeType, OrderType, ProductType, _
            Quantity, Price, TriggerPrice, Target, Stoploss, _
            TrailingStoploss, DisclosedQuantity, Validity, _
            Amo, StrategyId, Comments)

End Function

Public Function CancelOrder(pseudoAccount As String, _
        orderId As String) As Boolean

    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 2) As String

    cols(0) = CANCEL_ORDER_CMD
    cols(1) = pseudoAccount
    cols(2) = orderId

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

Public Function CancelOrderChildren(pseudoAccount As String, _
        orderId As String) As Boolean

    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 2) As String

    cols(0) = CANCEL_CHILD_ORDER_CMD
    cols(1) = pseudoAccount
    cols(2) = orderId

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

Public Function CancelAllOrders(pseudoAccount As String) As Boolean

    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 2) As String

    cols(0) = CANCEL_ALL_ORDERS_CMD
    cols(1) = pseudoAccount

    csv = Join(cols, ",")
        
    WriteCommand (csv)
        
    CancelAllOrders = True

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

Public Function ModifyOrder(pseudoAccount As String, _
    orderId As String, _
    OrderType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double) As Boolean
        
    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 6) As String

    cols(0) = MODIFY_ORDER_CMD
    cols(1) = pseudoAccount
    cols(2) = orderId
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

Public Function ModifyOrderPrice(pseudoAccount As String, _
    orderId As String, _
    Price As Double) As Boolean
        
    ModifyOrderPrice = ModifyOrder(pseudoAccount, orderId, "", 0, Price, 0)

End Function

Public Function ModifyOrderQuantity(pseudoAccount As String, _
    orderId As String, _
    Quantity As Long) As Boolean

    ModifyOrderQuantity = ModifyOrder(pseudoAccount, orderId, "", Quantity, 0, 0)

End Function

' Submits a square-off request for the given position.
'
' pseudoAccount - account to which the position belongs
' category - position category (DAY, NET). Pass DAY if you are not sure.
' type - position type (MIS, NRML, CNC, BO, CO)
' independentExchange - broker independent exchange
' independentSymbol - broker independent symbol
Public Function SquareOffPosition(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Boolean
        
    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 5) As String

    cols(0) = SQUARE_OFF_POSITION_CMD
    cols(1) = pseudoAccount
    cols(2) = category
    cols(3) = posType
    cols(4) = independentExchange
    cols(5) = independentSymbol

    csv = Join(cols, ",")
        
    WriteCommand (csv)
        
    SquareOffPosition = True

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

' Submits a square-off request for the given account.
' Server will square-off all open positions in the given account.
'
' pseudoAccount - account to which the position belongs
' category - position category (DAY, NET). Pass DAY if you are not sure.
Public Function SquareOffPortfolio(pseudoAccount As String, _
    category As String) As Boolean
        
    On Error GoTo Error_Handler
        
    Dim csv As String
    Dim cols(0 To 2) As String

    cols(0) = SQUARE_OFF_PORTFOLIO_CMD
    cols(1) = pseudoAccount
    cols(2) = category
    
    csv = Join(cols, ",")
        
    WriteCommand (csv)
        
    SquareOffPortfolio = True

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

Public Function isAutoTraderClientMonitoring() As Boolean
    
    Dim CommandsFilePath As String
    
    CommandsFilePath = GetCommandsFilePath()
    isAutoTraderClientMonitoring = ValidateFile(CommandsFilePath, _
        "AutoTrader client is not monitoring commands file." & CommandsFilePath)

End Function

' *****************************************************************************
' ************************ ORDER DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads orders file and returns a column value for the given order id.
Public Function ReadOrderColumn(pseudoAccount As String, _
    orderId As String, columnIndex As Integer) As String
    Dim FilePath As String
    FilePath = GetPortfolioOrdersFile(pseudoAccount)
    ReadOrderColumn = FileReadCsvColumnByRowId(FilePath, orderId, 3, columnIndex)
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

' Retrieve order's product type (INTRADAY, DELIVERY, NORMAL, MTF).
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

' Checks whether order is open.
Public Function IsOrderOpen(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderOpen = (UCase(oStatus) = "OPEN" Or UCase(oStatus) = "TRIGGER_PENDING")
End Function

' Checks whether order is complete.
Public Function IsOrderComplete(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderComplete = UCase(oStatus) = "COMPLETE"
End Function

' Checks whether order is rejected.
Public Function IsOrderRejected(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderRejected = UCase(oStatus) = "REJECTED"
End Function

' Checks whether order is cancelled.
Public Function IsOrderCancelled(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderCancelled = UCase(oStatus) = "CANCELLED"
End Function

' *****************************************************************************
' ************************ ORDER DETAIL FUNCTIONS - END ***********************
' *****************************************************************************


' *****************************************************************************
' ************************ POSITION DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads positions file and returns a column value for the given position id.
' Position id is a combination of category, type, independentExchange & independentSymbol.
Public Function ReadPositionColumnInternal(pseudoAccount As String, _
    category As String, categoryColumnIndex As Integer, _
    posType As String, typeColumnIndex As Integer, _
    independentExchange As String, independentExchangeColumnIndex As Integer, _
    independentSymbol As String, independentSymbolColumnIndex As Integer, _
    columnIndex As Integer) As String
    
    On Error GoTo Done
    
    Dim temp As String
    Dim cols() As String
    Dim FilePath As String
    
    FilePath = GetPortfolioPositionsFile(pseudoAccount)
    ReadPositionColumnInternal = ""
    
    Open FilePath For Input As #1
    
    Do Until EOF(1)
        Line Input #1, temp
        cols = Split(temp, ",")
        
        If (cols(categoryColumnIndex - 1) = category) And _
            (cols(typeColumnIndex - 1) = posType) And _
            (cols(independentExchangeColumnIndex - 1) = independentExchange) And _
            (cols(independentSymbolColumnIndex - 1) = independentSymbol) _
        Then
            ReadPositionColumnInternal = cols(columnIndex - 1)
            Exit Do
        End If
            
    Loop
    
    Close #1

Done:
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

' Reads positions file and returns a column value for the given position id.
' Position id is a combination of category, type, independentExchange & independentSymbol.
Public Function ReadPositionColumn(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String, columnIndex As Integer) As String
    
    ReadPositionColumn = ReadPositionColumnInternal(pseudoAccount, _
        category, 4, posType, 3, independentExchange, 5, independentSymbol, 6, columnIndex)
End Function

' Retrieve positions's trading account.
Public Function GetPositionTradingAccount(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionTradingAccount = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 2)
End Function

' Retrieve positions's MTM (Mtm calculated by your stock broker).
Public Function GetPositionMtm(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionMtm = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 7))
End Function

' Retrieve positions's PNL (Pnl calculated by your stock broker).
Public Function GetPositionPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionPnl = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 8))
End Function

' Retrieve positions's AT PNL (Pnl calculated by AutoTrader Web).
Public Function GetPositionAtPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionAtPnl = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 31))
End Function

' Retrieve positions's buy quantity.
Public Function GetPositionBuyQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionBuyQuantity = CLng(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 9))
End Function

' Retrieve positions's sell quantity.
Public Function GetPositionSellQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionSellQuantity = CLng(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 10))
End Function

' Retrieve positions's net quantity.
Public Function GetPositionNetQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionNetQuantity = CLng(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 11))
End Function

' Retrieve positions's buy value.
Public Function GetPositionBuyValue(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionBuyValue = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 12))
End Function

' Retrieve positions's sell value.
Public Function GetPositionSellValue(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionSellValue = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 13))
End Function

' Retrieve positions's net value.
Public Function GetPositionNetValue(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionNetValue = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 14))
End Function

' Retrieve positions's buy average price.
Public Function GetPositionBuyAvgPrice(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionBuyAvgPrice = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 15))
End Function

' Retrieve positions's sell average price.
Public Function GetPositionSellAvgPrice(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionSellAvgPrice = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 16))
End Function

' Retrieve positions's realised pnl.
Public Function GetPositionRealisedPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionRealisedPnl = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 17))
End Function

' Retrieve positions's unrealised pnl.
Public Function GetPositionUnrealisedPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionUnrealisedPnl = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 18))
End Function

' Retrieve positions's overnight quantity.
Public Function GetPositionOvernightQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionOvernightQuantity = CLng(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 19))
End Function

' Retrieve positions's multiplier.
Public Function GetPositionMultiplier(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionMultiplier = CLng(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 20))
End Function

' Retrieve positions's LTP.
Public Function GetPositionLtp(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionLtp = CDbl(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 21))
End Function

' Retrieve positions's (platform specific) exchange.
Public Function GetPositionExchange(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionExchange = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 22)
End Function

' Retrieve positions's (platform specific) symbol.
Public Function GetPositionSymbol(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionSymbol = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 23)
End Function

' Retrieve positions's date (DD-MM-YYYY).
Public Function GetPositionDay(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionDay = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 24)
End Function

' Retrieve positions's trading platform.
Public Function GetPositionPlatform(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionPlatform = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 25)
End Function

' Retrieve positions's account id as received from trading platform.
Public Function GetPositionAccountId(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionAccountId = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 26)
End Function

' Retrieve positions's stock broker.
Public Function GetPositionStockBroker(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionStockBroker = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 28)
End Function

' *****************************************************************************
' ************************ POSITION DETAIL FUNCTIONS - END ***********************
' *****************************************************************************


' *****************************************************************************
' ************************ MARGIN DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads margins file and returns a column value for the given margin category.
Public Function ReadMarginColumn(pseudoAccount As String, _
    category As String, columnIndex As Integer) As String
    Dim FilePath As String
    FilePath = GetPortfolioMarginsFile(pseudoAccount)
    ReadMarginColumn = FileReadCsvColumnByRowId(FilePath, category, 3, columnIndex)
End Function

' Retrieve margin funds.
Public Function GetMarginFunds(pseudoAccount As String, _
    category As String) As Double
    GetMarginFunds = CDbl(ReadMarginColumn(pseudoAccount, category, 4))
End Function

' Retrieve margin utilized.
Public Function GetMarginUtilized(pseudoAccount As String, _
    category As String) As Double
    GetMarginUtilized = CDbl(ReadMarginColumn(pseudoAccount, category, 5))
End Function

' Retrieve margin available.
Public Function GetMarginAvailable(pseudoAccount As String, _
    category As String) As Double
    GetMarginAvailable = CDbl(ReadMarginColumn(pseudoAccount, category, 6))
End Function

' Retrieve margin funds for equity category.
Public Function GetMarginFundsEquity(pseudoAccount As String) As Double
    GetMarginFundsEquity = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_EQUITY, 4))
End Function

' Retrieve margin utilized for equity category.
Public Function GetMarginUtilizedEquity(pseudoAccount As String) As Double
    GetMarginUtilizedEquity = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_EQUITY, 5))
End Function

' Retrieve margin available for equity category.
Public Function GetMarginAvailableEquity(pseudoAccount As String) As Double
    GetMarginAvailableEquity = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_EQUITY, 6))
End Function

' Retrieve margin funds for commodity category.
Public Function GetMarginFundsCommodity(pseudoAccount As String) As Double
    GetMarginFundsCommodity = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_COMMODITY, 4))
End Function

' Retrieve margin utilized for commodity category.
Public Function GetMarginUtilizedCommodity(pseudoAccount As String) As Double
    GetMarginUtilizedCommodity = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_COMMODITY, 5))
End Function

' Retrieve margin available for commodity category.
Public Function GetMarginAvailableCommodity(pseudoAccount As String) As Double
    GetMarginAvailableCommodity = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_COMMODITY, 6))
End Function

' Retrieve margin funds for entire account.
Public Function GetMarginFundsAll(pseudoAccount As String) As Double
    GetMarginFundsAll = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_ALL, 4))
End Function

' Retrieve margin utilized for entire account.
Public Function GetMarginUtilizedAll(pseudoAccount As String) As Double
    GetMarginUtilizedAll = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_ALL, 5))
End Function

' Retrieve margin available for entire account.
Public Function GetMarginAvailableAll(pseudoAccount As String) As Double
    GetMarginAvailableAll = CDbl(ReadMarginColumn(pseudoAccount, MARGIN_ALL, 6))
End Function

' *****************************************************************************
' ************************ MARGIN DETAIL FUNCTIONS - END ***********************
' *****************************************************************************

' *****************************************************************************
' ************************ PORTFOLIO SUMMARY FUNCTIONS - START ******************
' *****************************************************************************

' Reads summary file and returns a column value.
Public Function ReadSummaryColumn(pseudoAccount As String, _
    columnIndex As Integer) As String
    Dim FilePath As String
    FilePath = GetPortfolioSummaryFile(pseudoAccount)
    ReadSummaryColumn = FileReadCsvColumnByRowId(FilePath, pseudoAccount, 1, columnIndex)
End Function

' Retrieve portfolio M2M (Position category = DAY).
Public Function GetPortfolioMtm(pseudoAccount As String) As Double
    GetPortfolioMtm = CDbl(ReadSummaryColumn(pseudoAccount, 2))
End Function

' Retrieve portfolio PNL (Position category = DAY).
Public Function GetPortfolioPnl(pseudoAccount As String) As Double
    GetPortfolioPnl = CDbl(ReadSummaryColumn(pseudoAccount, 3))
End Function

' Retrieve portfolio position count (Position category = DAY).
Public Function GetPortfolioPositionCount(pseudoAccount As String) As Long
    GetPortfolioPositionCount = CLng(ReadSummaryColumn(pseudoAccount, 4))
End Function

' Retrieve portfolio OPEN position count (Position category = DAY).
Public Function GetPortfolioOpenPositionCount(pseudoAccount As String) As Long
    GetPortfolioOpenPositionCount = CLng(ReadSummaryColumn(pseudoAccount, 5))
End Function

' Retrieve portfolio CLOSED position count (Position category = DAY).
Public Function GetPortfolioClosedPositionCount(pseudoAccount As String) As Long
    GetPortfolioClosedPositionCount = CLng(ReadSummaryColumn(pseudoAccount, 6))
End Function

' Retrieve portfolio open short quantity (Position category = DAY).
Public Function GetPortfolioOpenShortQuantity(pseudoAccount As String) As Long
    GetPortfolioOpenShortQuantity = CLng(ReadSummaryColumn(pseudoAccount, 7))
End Function

' Retrieve portfolio open long quantity (Position category = DAY).
Public Function GetPortfolioOpenLongQuantity(pseudoAccount As String) As Long
    GetPortfolioOpenLongQuantity = CLng(ReadSummaryColumn(pseudoAccount, 8))
End Function

' Retrieve portfolio order count.
Public Function GetPortfolioOrderCount(pseudoAccount As String) As Long
    GetPortfolioOrderCount = CLng(ReadSummaryColumn(pseudoAccount, 9))
End Function

' Retrieve portfolio "open" order count.
Public Function GetPortfolioOpenOrderCount(pseudoAccount As String) As Long
    GetPortfolioOpenOrderCount = CLng(ReadSummaryColumn(pseudoAccount, 10))
End Function

' Retrieve portfolio "complete" order count.
Public Function GetPortfolioCompleteOrderCount(pseudoAccount As String) As Long
    GetPortfolioCompleteOrderCount = CLng(ReadSummaryColumn(pseudoAccount, 11))
End Function

' Retrieve portfolio "cancelled" order count.
Public Function GetPortfolioCancelledOrderCount(pseudoAccount As String) As Long
    GetPortfolioCancelledOrderCount = CLng(ReadSummaryColumn(pseudoAccount, 12))
End Function

' Retrieve portfolio "rejected" order count.
Public Function GetPortfolioRejectedOrderCount(pseudoAccount As String) As Long
    GetPortfolioRejectedOrderCount = CLng(ReadSummaryColumn(pseudoAccount, 13))
End Function

' Retrieve portfolio "trigger pending" order count.
Public Function GetPortfolioTriggerPendingOrderCount(pseudoAccount As String) As Long
    GetPortfolioTriggerPendingOrderCount = CLng(ReadSummaryColumn(pseudoAccount, 14))
End Function

' Retrieve portfolio M2M (Position category = NET).
Public Function GetPortfolioMtmNET(pseudoAccount As String) As Double
    GetPortfolioMtmNET = CDbl(ReadSummaryColumn(pseudoAccount, 15))
End Function

' Retrieve portfolio PNL (Position category = NET).
Public Function GetPortfolioPnlNET(pseudoAccount As String) As Double
    GetPortfolioPnlNET = CDbl(ReadSummaryColumn(pseudoAccount, 16))
End Function

' Retrieve portfolio position count (Position category = NET).
Public Function GetPortfolioPositionCountNET(pseudoAccount As String) As Long
    GetPortfolioPositionCountNET = CLng(ReadSummaryColumn(pseudoAccount, 17))
End Function

' Retrieve portfolio OPEN position count (Position category = NET).
Public Function GetPortfolioOpenPositionCountNET(pseudoAccount As String) As Long
    GetPortfolioOpenPositionCountNET = CLng(ReadSummaryColumn(pseudoAccount, 18))
End Function

' Retrieve portfolio CLOSED position count (Position category = NET).
Public Function GetPortfolioClosedPositionCountNET(pseudoAccount As String) As Long
    GetPortfolioClosedPositionCountNET = CLng(ReadSummaryColumn(pseudoAccount, 19))
End Function

' Retrieve portfolio open short quantity (Position category = NET).
Public Function GetPortfolioOpenShortQuantityNET(pseudoAccount As String) As Long
    GetPortfolioOpenShortQuantityNET = CLng(ReadSummaryColumn(pseudoAccount, 20))
End Function

' Retrieve portfolio open long quantity (Position category = NET).
Public Function GetPortfolioOpenLongQuantityNET(pseudoAccount As String) As Long
    GetPortfolioOpenLongQuantityNET = CLng(ReadSummaryColumn(pseudoAccount, 21))
End Function

' *****************************************************************************
' ************************ PORTFOLIO SUMMARY FUNCTIONS - END ********************
' *****************************************************************************

