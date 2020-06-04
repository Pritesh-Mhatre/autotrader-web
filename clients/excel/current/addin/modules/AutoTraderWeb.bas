Attribute VB_Name = "AutoTraderWeb"
Option Explicit

Dim ORDER_NUM As Integer
Dim START_TIME As Long

Const COMMANDS_FILE As String = "commands.csv"

Const INPUT_DIR As String = "input"
Const OUTPUT_DIR As String = "output"

Const CANCEL_ORDER_CMD As String = "CANCEL_ORDER"
Const MODIFY_ORDER_CMD As String = "MODIFY_ORDER"

Public Const EPOCH As Date = #1/1/1970#
Public Const BLANK As String = ""

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

Private Function GetIPCDirectory() As String
    
    GetIPCDirectory = Environ("USERPROFILE") & Application.PathSeparator & "autotrader"

End Function

Private Function GetCommandsFilePath() As String
    
    GetCommandsFilePath = GetIPCDirectory & Application.PathSeparator _
        & INPUT_DIR & Application.PathSeparator & COMMANDS_FILE

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
        Message = str(Err.Number) & Err.Description
        MsgBox Message, vbOKOnly, "Error"
        Resume Next

    End If
    
End Function

Public Function PlaceOrder(Variety As String, _
    PseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    ProductType As String, _
    OrderType As String, _
    Quantity As Integer, _
    Price As Double, _
    TriggerPrice As Double, _
    Amo As Boolean, _
    Validity As String, _
    DisclosedQuantity As Integer, _
    PublisherId As String, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double, _
    Comments As String, _
    StrategyId As Integer) As String

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
    o.PublisherId = PublisherId
    o.Target = Target
    o.Stoploss = Stoploss
    o.TrailingStoploss = TrailingStoploss
    o.Comments = Comments
    o.StrategyId = StrategyId

    PlaceOrder = PlaceOrderInternal(o)

End Function

Public Function isAutoTraderClientMonitoring() As Boolean
    
    Dim CommandsFilePath As String
    
    CommandsFilePath = GetCommandsFilePath()
    isAutoTraderClientMonitoring = ValidateFile(CommandsFilePath, _
        "AutoTrader client is not monitoring commands file.")

End Function