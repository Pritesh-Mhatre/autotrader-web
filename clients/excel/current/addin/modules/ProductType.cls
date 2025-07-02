VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum ProductTypeEnum
    INTRADAY
    DELIVERY
    NORMAL
	MTF
End Enum

Private Const LoLimit = ProductTypeEnum.INTRADAY
Private Const HiLimit = ProductTypeEnum.MTF

Private gen As ProductTypeEnum

Private Sub Class_Initialize()
    gen = ProductTypeEnum.INTRADAY
End Sub

Public Property Get Value() As ProductTypeEnum
Attribute Value.VB_UserMemId = 0
    Value = gen
End Property

Public Property Let Value(val As ProductTypeEnum)
    If val >= LoLimit And val <= HiLimit Then
        gen = val
    Else
        Err.Raise Number:=vbObjectError + 513, Description:="Invalid value for ProductType"
    End If
End Property

Public Sub FromString(str As String)
    Select Case Trim(UCase(str))
        Case "I", "MIS", "INTRADAY": Me.Value = ProductTypeEnum.INTRADAY
                Case "D", "CNC", "DELIVERY": Me.Value = ProductTypeEnum.DELIVERY
                Case "NRML", "NORMAL": Me.Value = ProductTypeEnum.NORMAL
                Case "MTF": Me.Value = ProductTypeEnum.MTF
        Case Else: Me.Value = ProductTypeEnum.INTRADAY
    End Select
End Sub

Public Function ToString() As String
    Dim str As String

    Select Case gen
        Case ProductTypeEnum.INTRADAY: str = "INTRADAY"
        Case ProductTypeEnum.DELIVERY: str = "DELIVERY"
        Case ProductTypeEnum.NORMAL: str = "NORMAL"
        Case ProductTypeEnum.MTF: str = "MTF"
        Case Else: str = "INTRADAY"
    End Select
    ToString = str
End Function
