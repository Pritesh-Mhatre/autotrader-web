VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum ValidityEnum
    DAY
    IOC
End Enum

Private Const LoLimit = ValidityEnum.DAY
Private Const HiLimit = ValidityEnum.IOC

Private gen As ValidityEnum

Private Sub Class_Initialize()
    gen = ValidityEnum.DAY
End Sub

Public Property Get Value() As ValidityEnum
Attribute Value.VB_UserMemId = 0
    Value = gen
End Property

Public Property Let Value(val As ValidityEnum)
    If val >= LoLimit And val <= HiLimit Then
        gen = val
    Else
        Err.Raise Number:=vbObjectError + 513, Description:="Invalid value for Validity"
    End If
End Property

Public Sub FromString(str As String)
    Select Case Trim(UPPER(str))
        Case "IOC": Me.Value = ValidityEnum.IOC
        Case Else: Me.Value = ValidityEnum.DAY
    End Select
End Sub

Public Function ToString() As String
    Dim str As String

    Select Case gen
        Case ValidityEnum.IOC: str = "IOC"
        Case Else: str = "DAY"
    End Select
    ToString = str
End Function
