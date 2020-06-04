VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PLACE_ORDER_CMD As String = "PLACE_ORDER"

Public Variety As Variety
Public PseudoAccount As String
Public Exchange As String
Public Symbol As String
Public TradeType As TradeType
Public OrderType As OrderType
Public ProductType As ProductType
Public Quantity As Integer
Public Price As Double
Public TriggerPrice As Double
Public Amo As Boolean
Public Validity As Validity
Public DisclosedQuantity As Integer
Public PublisherId As String
Public Target As Double
Public Stoploss As Double
Public TrailingStoploss As Double
Public Comments As String
Public StrategyId As Integer
Public PublishTime As Date


Private Sub Class_Initialize()
    
    Set Me.Variety = New Variety
    Set Me.Amo = False
    Set Me.Validity = New Validity
    Set Me.DisclosedQuantity = 0
    Set Me.StrategyId = 0
    Set Me.Comments = ""
    Set Me.TriggerPrice = 0
    Set Me.Price = 0
    Set Me.Target = 0
    Set Me.Stoploss = 0
    Set Me.TrailingStoploss = 0
    Set Me.PublishTime = Now

End Sub

Public Function ToPlaceCommand() As String
    
    Dim csv As String
    Dim cols(0 To 20) As String

    cols(0) = PLACE_ORDER_CMD
    cols(1) = Me.PseudoAccount
    cols(2) = Me.PublisherId
    cols(3) = Me.Variety.ToString
    cols(4) = Me.Exchange
    cols(5) = Me.Symbol
    cols(6) = Me.TradeType.ToString
    cols(7) = Me.OrderType.ToString
    cols(8) = Me.ProductType.ToString
    cols(9) = Me.Quantity
    cols(10) = Me.Price
    cols(11) = Me.TriggerPrice
    cols(12) = Me.Target
    cols(13) = Me.Stoploss
    cols(14) = Me.TrailingStoploss
    cols(15) = Me.DisclosedQuantity
    cols(16) = Me.Validity.ToString
    cols(17) = Me.Amo
    ' TODO Add time as millis since epoch
    cols(18) = CStr((Me.PublishTime - EPOCH) * 86400)
    cols(19) = Me.StrategyId
    cols(20) = Me.Comments
    
    csv = Join(cols, ",")
    
    ToPlaceCommand = csv

End Function
