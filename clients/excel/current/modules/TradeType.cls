VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TradeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum TradeTypeEnum
    BUY
    SELL
End Enum

Private Const LoLimit = TradeTypeEnum.BUY
Private Const HiLimit = TradeTypeEnum.SELL

Private gen As TradeTypeEnum

Private Sub Class_Initialize()
    gen = TradeTypeEnum.BUY
End Sub

Public Property Get Value() As TradeTypeEnum
Attribute Value.VB_UserMemId = 0
    Value = gen
End Property

Public Property Let Value(val As TradeTypeEnum)
    If val >= LoLimit And val <= HiLimit Then
        gen = val
    Else
        Err.Raise Number:=vbObjectError + 513, Description:="Invalid value for TradeType"
    End If
End Property

Public Sub FromString(str As String)
    Select Case Trim(UPPER(str))
        Case "S", "SELL": Me.Value = TradeTypeEnum.SELL
        Case Else: Me.Value = TradeTypeEnum.BUY
    End Select
End Sub

Public Function ToString() As String
    Dim str As String

    Select Case gen
        Case TradeTypeEnum.SELL: str = "SELL"
        Case Else: str = "BUY"
    End Select
    ToString = str
End Function
