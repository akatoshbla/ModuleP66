Attribute VB_Name = "P66Module"
Private Const BROADCAST_NETID = &H400000
Private Const PCANHANDLE = PCANBasic.PCAN_USBBUS1

Private mm_enabled As Boolean
Private canbus_Connected As Boolean
Private netID As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

''' <summary>
''' Initializes a PCAN Channel
''' </summary>
''' <returns>"A boolean: True means that the initialization was successful, False means that the Initialization failed."</returns>
Public Function Canbus_Initialize() As Boolean
    Dim result As Long
    Dim ibuffer As Integer
    Dim IOPort As Integer
    Dim baudrate As Long
    Dim HwType As Long
    
    mm_enabled = False
    canbus_Connected = False
    
    ibuffer = 1
    IOPort = 256
    baudrate = PCANBasic.PCAN_BAUD_250K
    HwType = PCANBasic.PCAN_TYPE_ISA
    
    result = PCANBasic.CAN_Initialize(PCANHANDLE, baudrate, HwType, IOPort, CByte("&H" & 3))
    
    If result = PCANBasic.PCAN_ERROR_OK Then
        canbus_Connected = True
        Canbus_Initialize = True
    Else
        canbus_Connected = False
        Canbus_Initialize = False
    End If
End Function

''' <summary>
''' Uninitializes all PCAN Channels initialized by Canbus_Initialize
''' </summary>
Public Sub Canbus_Uninitialize()
    mm_enabled = False
    canbusConnected = False
    PCANBasic.CAN_Uninitialize (PCANBasic.PCAN_USBBUS1)
End Sub

''' <summary>
'''
''' </summary>
''' <param name="command">"The command for the transducer"</param>
''' <param name="serialNumber">"Optional: SN of transducer"</param>
''' <param name="net_ID">"Optional: Net ID of transducer"</param>
''' <param name="value">"Optional: Dac Value that is to be written to the Transducer"</param>
''' <returns>"Parsed string of the transducer's output"</returns>
Public Function Canbus_Send(command As String, Optional serialNumber As String, Optional net_ID As String, Optional value As String) As String
    Dim DATA() As Long
    Dim writeResult As Long
    Dim result As String
    
    If mm_enabled = False Then
        Canbus_Send = "Error: Not in MM"
    End If
    
    ' TODO: Error control for checking connection is alive
    
    Select Case command
        Case "MM"
            DATA = Canbus_DataPacket(command, False)
            writeResult = Canbus_Write(DATA, False)
            result = Canbus_Read(command, 4, writeResult)
        Case "p", "t", "d1", "d2", "d3"
            DATA = Canbus_DataPacket(command, False)
            writeResult = Canbus_Write(DATA, False, net_ID)
            result = Canbus_Read(command, 8, writeResult)
        Case "I"
            DATA = Canbus_DataPacket(command, True, serialNumber)
            writeResult = Canbus_Write(DATA, False, net_ID)
            result = Canbus_Read(command, 8, writeResult)
        Case "D1", "D2", "D3"
            DATA = Canbus_DataPacket(command, True, "", value)
            writeResult = Canbus_Write(DATA, True, net_ID)
            result = Canbus_Read(command, 8, writeResult)
        Case Else
            Canbus_Send = "Send Error: Invalid Command"
    End Select
    
    Canbus_Send = result
End Function

''' This function makes the Data Packet for the Canbus MSG - Returns a Long Array
Private Function Canbus_DataPacket(command As String, writing As Boolean, Optional serialNumber As String, Optional value As String) As Long()
    Dim msg_StrArr() As String
    Dim msg() As String
    Dim DATA(7) As Long
    
    If writing = False Then
        msg_StrArr = str2Array(command)
        ReDim msg(UBound(msg_StrArr))
        For i = 0 To UBound(msg_StrArr)
            msg(i) = string2Hex(msg_StrArr(i))
        Next i
    Else
        Dim temp As String
        
        msg_StrArr = str2Array(command)
        For i = 0 To UBound(msg_StrArr)
            temp = temp & Hex(Asc(msg_StrArr(i)))
        Next i
        
        If Len(serialNumber) = 6 Then
            temp = temp & littleEndian(Hex(Int(serialNumber))) & littleEndian(Hex(createNetID(serialNumber)))
        Else
            temp = temp & littleEndian(Format(Hex(CLng(value)), "0000"))
        End If
        
        msg = str2ArrayPair(temp)
    End If
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        DATA(i) = CByte("&H" & msg(i))
    Next i
    
    Canbus_DataPacket = DATA
End Function

''' This Function writes to the packet to the Canbus - Returns a Long which is the result of the write. 0 means no error
Private Function Canbus_Write(DATA() As Long, writing As Boolean, Optional net_ID As String) As Long
    Dim CANMsg As TPCANMsg
    Dim result As Long
    Dim dataLen As Integer
    
    If net_ID <> "" Then
        CANMsg.ID = CLng(net_ID)
    Else
        CANMsg.ID = BROADCAST_NETID
    End If
    
    CANMsg.LEN = CByte(UBound(DATA) + 1)
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    For i = 0 To UBound(DATA) - 1
        CANMsg.DATA(i) = DATA(i)
    Next i
    
    If canbus_Connected = True Then
        result = PCANBasic.CAN_Write(PCANHANDLE, CANMsg)
    Else
        result = 999999
    End If
    
    Canbus_Write = result
End Function

''' This function reads the Canbus buffer and finds the correct packet to read. Returns the parsed value of the returning packet.
Private Function Canbus_Read(command As String, length As Integer, writeResult As Long) As String
    Dim CANMsg As TPCANMsg
    Dim timeStamp As TPCANTimestamp
    Dim dataLength As Byte
    Dim result As Long
    Dim commandByte As String
    Dim strResult As String
    
    dataLength = CByte(length)
    strResult = ""
    
    Sleep (500)
    
    If writeResult <> PCANBasic.PCAN_ERROR_OK Then
        Canbus_Read = "Error: Writing to CANBUS failed check Canbus_Write method : " & writeResult
    End If
    
    result = PCANBasic.CAN_Read(PCANHANDLE, CANMsg, timeStamp)
    
    If result = PCANBasic.PCAN_ERROR_OK Then
        While CANMsg.LEN <> dataLength
            result = PCANBasic.CAN_Read(PCANHANDLE, CANMsg, timeStamp)
        Wend
    Else
        Canbus_Read = "Error: Reading Buffer : " & result
    End If
    
    commandByte = Chr(CANMsg.DATA(0))
    
    For i = 0 To CANMsg.LEN - 1
        strResult = strResult & Format(Hex(CANMsg.DATA(i)), "00")
    Next i
    
    If commandByte = Trim(Mid(command, 1, 1)) Then
        If commandByte = "M" Then
            Canbus_Read = "MME"
        ElseIf commandByte = "I" Then
            Canbus_Read = "Net_ID is set"
        ElseIf Len(command) = 1 Then
            Canbus_Read = CLng("&H" & littleEndian(Trim(Mid(strResult, 3, 4))))
        ElseIf Len(command) = 2 Then
            Canbus_Read = CLng("&H" & littleEndian(Trim(Mid(strResult, 5, 4))))
        End If
    Else
        Canbus_Read = "Error: Could not find Command result in CANBUS Buffer"
    End If
    
End Function

''' This function helps set the transducer's netID returns the OR of the broadcast and serialNumber
Private Function createNetID(serialNumber As String) As Long
    netID = BROADCAST_NETID Or CLng(serialNumber)
    createNetID = netID
End Function

''' Helper function to convert Hex to binary
Private Function Hex2Bin(HexStr As String) As Byte()
    ReDim Bytes(0 To Len(HexStr) \ 2 - 1) As Byte
    Dim i As Long
    Dim j As Long
    For i = 1 To Len(HexStr) Step 2
        Bytes(j) = CByte("&H" & Mid$(HexStr, i, 2))
        j = j + 1
        Next
    HexBytes = Bytes
End Function

''' Helper function that converts a string to a hex value
Private Function string2Hex(xString As String) As String
For i = 1 To Len(xString)
    string2Hex = string2Hex & Hex$(Asc(Mid(xString, i, 1))) & Space$(1)
    On Error Resume Next
    DoEvents
    Next i
    string2Hex = Mid$(string2Hex, 1, Len(string2Hex) - 1)
End Function

''' Helper function that converts a string to a char array
Private Function str2Array(xString As String) As String()
Dim tempArray() As String
ReDim tempArray(Len(xString) - 1)

For i = 1 To Len(xString)
    tempArray(i - 1) = Mid(xString, i, 1)
Next i

str2Array = tempArray
End Function

''' Helper Function that converts a string into a paired array. I.E. "1234" would be "(12)(34)"
Private Function str2ArrayPair(xString As String) As String()
    Dim tempArray() As String
    ReDim tempArray((Len(xString) / 2) - 1)
    Dim k As Integer
    
    k = 3
    
    For i = 1 To UBound(tempArray) + 1
        If i = 1 Then
            tempArray(i - 1) = Mid$(xString, i, 2)
        Else
            tempArray(i - 1) = Mid$(xString, k, 2)
            k = k + 2
        End If
    Next i

str2ArrayPair = tempArray
End Function

''' Helper function that takes a string and reverses its pairs. I.E. "F1F2F3" would be "F3F2F1"
Private Function littleEndian(bigEndian As String) As String
    Dim temp() As String
    Dim result As String
    Dim length As Integer
    
    result = ""
    
    If Len(bigEndian) Mod 2 = 1 Then
        bigEndian = "0" & bigEndian
    End If
    
    temp = str2Array(bigEndian)
    length = UBound(temp)
      
    For i = length To 0 Step -2
            result = result & temp(i - 1) & temp(i)
        Next i
    
    littleEndian = result
    
End Function
