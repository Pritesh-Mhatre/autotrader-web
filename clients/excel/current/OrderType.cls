VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum OrderTypeEnum
    LIMIT
    MARKET
        STOP_LOSS
        SL_MARKET
End Enum

Private Const LoLimit = OrderTypeEnum.LIMIT
Private Const HiLimit = OrderTypeEnum.SL_MARKET

Private gen As OrderTypeEnum

Private Sub Class_Initialize()
    gen = OrderTypeEnum.LIMIT
End Sub

Public Property Get Value() As OrderTypeEnum
Attribute Value.VB_UserMemId = 0
    Value = gen
End Property

Public Property Let Value(val As OrderTypeEnum)
    If val >= LoLimit And val <= HiLimit Then
        gen = val
    Else
        Err.Raise Number:=vbObjectError + 513, Description:="Invalid value for OrderType"
    End If
End Property

Public Sub FromString(str As String)
    Select Case Trim(UPPER(str))
        Case "M", "MARKET", "MKT": Me.Value = OrderTypeEnum.MARKET
                Case "SL", "STOP_LOSS": Me.Value = OrderTypeEnum.STOP_LOSS
                Case "SL-M", "SL_MARKET": Me.Value = OrderTypeEnum.SL_MARKET
        Case Else: Me.Value = OrderTypeEnum.LIMIT
    End Select
End Sub

Public Function ToString() As String
    Dim str As String

    Select Case gen
        Case OrderTypeEnum.MARKET: str = "MARKET"
        Case OrderTypeEnum.STOP_LOSS: str = "STOP_LOSS"
        Case OrderTypeEnum.SL_MARKET: str = "SL_MARKET"
        Case Else: str = "LIMIT"
    End Select
    ToString = str
End Function
