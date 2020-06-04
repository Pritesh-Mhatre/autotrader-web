VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Variety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum VarietyEnum
    REGULAR
    BO
    CO
End Enum

Private Const LoLimit = VarietyEnum.REGULAR
Private Const HiLimit = VarietyEnum.CO

Private gen As VarietyEnum

Private Sub Class_Initialize()
    gen = VarietyEnum.REGULAR
End Sub

Public Property Get Value() As VarietyEnum
Attribute Value.VB_UserMemId = 0
    Value = gen
End Property

Public Property Let Value(val As VarietyEnum)
    If val >= LoLimit And val <= HiLimit Then
        gen = val
    Else
        Err.Raise Number:=vbObjectError + 513, Description:="Invalid value for Variety"
    End If
End Property

Public Sub FromString(str As String)
    Select Case Trim(UPPER(str))
        Case "CO": Me.Value = VarietyEnum.CO
        Case "BO": Me.Value = VarietyEnum.BO
        Case Else: Me.Value = VarietyEnum.REGULAR
    End Select
End Sub

Public Function ToString() As String
    Dim str As String

    Select Case gen
        Case VarietyEnum.BO: str = "BO"
        Case VarietyEnum.CO: str = "CO"
        Case Else: str = "REGULAR"
    End Select
    ToString = str
End Function
