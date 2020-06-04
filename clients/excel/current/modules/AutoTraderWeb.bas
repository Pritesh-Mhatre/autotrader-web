Attribute VB_Name = "AutoTraderWeb"
Option Explicit

Dim ORDER_NUM As Integer
Dim START_TIME As Long

Const COMMANDS_FILE As String = "commands.csv"

Const CANCEL_ORDER_CMD As String = "CANCEL_ORDER"
Const MODIFY_ORDER_CMD As String = "MODIFY_ORDER"

Const EPOCH As Date = #1/1/1970#
Const BLANK As String = ""

Public Function NextOrderNumber() As String

    If ORDER_NUM = 0 Then
        START_TIME = Abs(CLng((Now() - EPOCH) * 86400 - 2 ^ 31))
    End If
    
    ORDER_NUM = ORDER_NUM + 1
    NextOrderNumber = CStr(START_TIME + ORDER_NUM)
    
End Function

Public Function ValidateFile(FilePath As String, Message As String) As Boolean

    With (CreateObject("Scripting.FileSystemObject"))
        If Not .FileExists(FilePath) Then
            MsgBox Message, vbOKOnly, "Error"
            ValidateFile = False
        Else
            ValidateFile = True
        End If
    End With

End Function

Public Function GetIPCDirectory() As String
    
    GetIPCDirectory = Environ("USERPROFILE") & Application.PathSeparator & "autotrader"

End Function

Public Function GetCommandsFilePath() As String
    
    GetCommandsFilePath = GetIPCDirectory & Application.PathSeparator & COMMANDS_FILE

End Function

Public Sub WriteCommand(Command As String)

    Dim CommandsFilePath As String
    
    CommandsFilePath = GetCommandsFilePath()
    If ValidateFile(CommandsFilePath, "AutoTrader client is not monitoring commands file.") = False Then
        Exit Sub
    End If

    Open CommandsFilePath For Append As #1
        Print #1, Command
    Close #1

End Sub

Public Function PlaceOrder(Order As Order) As String
    
    On Error GoTo Error_Handler
    
    ' Assign a unique order id
    Order.PublisherId = NextOrderNumber()
    
    ' Write PlaceOrder Command to File
    WriteCommand (Order.ToPlaceCommand)

    PlaceOrder = Order.PublisherId

Error_Handler:
    If Err.Number <> 0 Then
        
        Dim Message As String
        Message = str(Err.Number) & Err.Description
        MsgBox Message, vbOKOnly, "Error"
        Resume Next

    End If
    
End Function
