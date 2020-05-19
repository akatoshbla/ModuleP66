VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "P66 Utility"
   ClientHeight    =   6855
   ClientLeft      =   11070
   ClientTop       =   3900
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7160.969
   ScaleMode       =   0  'User
   ScaleWidth      =   7687.213
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Packet Sent"
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   2775
      Begin VB.TextBox tb_messageSent 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   275
         Width           =   2535
      End
   End
   Begin VB.TextBox tb_output 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Packet Details"
      Height          =   1935
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   4455
      Begin VB.TextBox tb_input 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "From P66"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "To P66"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox SSTab1 
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   7395
      TabIndex        =   4
      Top             =   2280
      Width           =   7455
      Begin VB.Frame Frame12 
         Caption         =   "Pressure Counts"
         Height          =   855
         Left            =   2640
         TabIndex        =   32
         Top             =   1680
         Width           =   2175
         Begin VB.TextBox tb_pressure 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Temp Counts"
         Height          =   855
         Left            =   2640
         TabIndex        =   30
         Top             =   840
         Width           =   2175
         Begin VB.TextBox tb_temp 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Counts"
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   2350
         Begin VB.CommandButton pressure_btn 
            Caption         =   "Pressure"
            Height          =   615
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton temp_btn 
            Caption         =   "Temperature"
            Height          =   615
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Span DAC"
         Height          =   855
         Left            =   4935
         TabIndex        =   23
         Top             =   0
         Width           =   2350
         Begin VB.TextBox tb_span 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2125
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Zero DAC"
         Height          =   855
         Left            =   2612
         TabIndex        =   22
         Top             =   0
         Width           =   2175
         Begin VB.TextBox tb_zero 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Gain POT"
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   2350
         Begin VB.TextBox tb_gain 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "One to One"
         Height          =   1815
         Left            =   2612
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
         Begin VB.CommandButton readSN_btn 
            Caption         =   "Read Serial Number"
            Height          =   615
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Network Protocol"
         Height          =   3495
         Left            =   4935
         TabIndex        =   8
         Top             =   840
         Width           =   2350
         Begin VB.CommandButton resetNID_btn 
            Caption         =   "Reset Node ID"
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   2520
            Width           =   1935
         End
         Begin VB.CommandButton readdacs_btn 
            Caption         =   "Read DACs"
            Height          =   615
            Left            =   240
            TabIndex        =   15
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton setdacs_btn 
            Caption         =   "Write to DACs"
            Height          =   615
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton setNID_btn 
            Caption         =   "Set Network ID"
            Height          =   615
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Connection"
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2350
         Begin VB.CommandButton endmm_btn 
            Caption         =   "Exit Maintenance Mode"
            Height          =   615
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton startmm_btn 
            Caption         =   "Enter Maintenance Mode"
            Height          =   615
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Sent to the P66"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.TextBox tb_netid 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   825
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Network ID"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_PcanHandle As Long
Dim m_Baudrate As Long
Dim m_HwType As Long
Dim s_netID As Long
Dim connected As Boolean
Dim mm_enabled As Boolean
Dim serialNumber As String
Dim readSNCalled As Boolean
Dim setNetIdCalled As Boolean
Dim busStatus As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub endmm_btn_Click()
    Call deInitCanbus
End Sub

Private Sub Form_Load()
    mm_enabled = False
    readSNCalled = False
    setNetIdCalled = False
    connected = False
    Call initCanbus
End Sub

Private Sub initCanbus()
    Dim stsResult As Long
    Dim ibuffer As Integer
    Dim IOPort As Integer
    
    ibuffer = 1
    IOPort = 256
    m_PcanHandle = PCANBasic.PCAN_USBBUS1
    m_Baudrate = PCANBasic.PCAN_BAUD_250K
    m_HwType = PCANBasic.PCAN_TYPE_ISA
    

    
    stsResult = PCANBasic.CAN_Initialize(m_PcanHandle, m_Baudrate, m_HwType, IOPort, CByte("&H" & 3))
    
    If stsResult <> PCANBasic.PCAN_ERROR_OK Then
        tb_input.Text = "Canbus Initialization Failed"
    Else
        tb_input.Text = "Canbus Initialized Successfully"
    End If
    
    stsResult = PCANBasic.CAN_SetValue(m_PcanHandle, PCANBasic.PCAN_BUSOFF_AUTORESET, ibuffer, CByte("&H" & 4))
    
    If stsResult <> PCANBasic.PCAN_ERROR_OK Then
        tb_output.Text = "Canbus Connection Failed!"
    Else
        connected = True
        tb_output.Text = "CANBUS Connected!"
    End If
End Sub

Private Sub deInitCanbus()
    mm_enabled = False
    readSNCalled = False
    connected = False
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    PCANBasic.CAN_Uninitialize (m_PcanHandle)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call deInitCanbus
End Sub

Private Sub pressure_btn_Click()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd As String
    Dim msg() As String
    Dim msg1() As String
    Dim dataLen As Integer
    
    cmnd = "p"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_pressure.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 3, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub readdacs_btn_Click()
    Call dacZero
    Call dacSpan
    Call dacGain
End Sub

Private Sub readSN_btn_Click()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd_r As String
    Dim msg() As String
    Dim msg1() As String
    Dim dataLen As Integer
    Dim netID As String
    
    netID = "00400000"
    cmnd_r = "R"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd_r)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = CLng("&H" & netID)
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    For i = 0 To UBound(CANMsg.DATA)
        If CANMsg.DATA(i) = CByte(0) Then
            CANMsg.DATA(i) = CByte("&H" & 0)
        End If
    Next i
    
    If connected = True And mm_enabled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True Then
            readSNCalled = True
            Call readMessages(7)
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub setdacs_btn_Click()
    Call writeZero
    Call writeSpan
    Call writeGain
End Sub

'Private Sub resetNID_btn_Click()
'    Dim CANMsg As TPCANMsg
'    Dim stsResult As Long
'    Dim cmnd As String
'    Dim msg() As String
'    Dim msg1() As String
'    Dim dataLen As Integer
'
'    cmnd = "K"
'
'    tb_input.Text = ""
'    tb_output.Text = ""
'    tb_messageSent.Text = ""
'    tb_messageSent.BackColor = vbWhite
'
'    msg1 = str2Array(cmnd)
'    ReDim msg(UBound(msg1))
'    For i = 0 To UBound(msg1)
'        msg(i) = string2Hex(msg1(i))
'    Next i
'
'    dataLen = UBound(msg1) + 1
'
'    CANMsg.ID = "00400000"
'
'    CANMsg.LEN = CByte(dataLen)
'
'    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
'
'    For i = 0 To UBound(msg)
'        If msg(i) = "" Then
'            msg(i) = "00"
'        End If
'        CANMsg.DATA(i) = CByte("&H" & msg(i))
'    Next i
'
'    If connected = True And mm_enabled = True Then
'        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
'    Else
'        stsResult = "999999"
'    End If
'
'    For i = 0 To UBound(CANMsg.DATA)
'        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
'    Next i
'
'    If stsResult = PCANBasic.PCAN_ERROR_OK Then
'        tb_messageSent.Text = "Message Sent Successfully"
'        tb_messageSent.BackColor = vbGreen
'        If connected = True And mm_enabled = True Then
'            Call readMessages(4)
'            '   tb_pressure.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 3, 4))))
'        End If
'    Else
'        tb_messageSent.Text = "Message Sent Failed"
'        tb_messageSent.BackColor = vbRed
'        tb_output.Text = stsResult
'    End If
'End Sub

Private Sub setNID_btn_Click()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim msg() As String
    Dim dataLen As Integer
    Dim packet As String
    Dim netID As String
    
    netID = "00400000"
    
    packet = Hex(Asc("I")) & littleEndian(Hex(Int(serialNumber))) & littleEndian(Hex(s_netID))
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg = str2ArrayPair(packet)
    
    dataLen = UBound(msg) + 1
    
    CANMsg.ID = CLng("&H" & netID)
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Hex(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True Then
            ' Switch to DACs tab
            Call readMessages(8)
            Call dacZero
            Call dacSpan
            Call dacGain
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub startmm_btn_Click()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd_mm As String
    Dim msg() As String
    Dim msg1() As String '  String Array
    Dim dataLen As Integer
    Dim netID As String
    
    netID = "400000"
    cmnd_mm = "MM"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd_mm)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = CLng("&H" & netID)
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
'    For i = 0 To UBound(CANMsg.DATA)
'        If CANMsg.DATA(i) = CByte(0) Then
'            CANMsg.DATA(i) = CByte("&H" & 0)
'        End If
'    Next i
    
    If connected = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = 999999
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        '   tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        mm_enabled = True
        If connected = True Then
            Call readMessages(4)
            stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
            Call readMessages(4)
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = "Entering Maintaince Mode Failed"
    End If
    
End Sub

Private Sub readMessages(packetLength As Integer)
    Dim CANMsg As TPCANMsg
    Dim CANTimeStamp As TPCANTimestamp
    Dim stsResult As Long
    Dim displayed As Boolean
    Dim dataLen As Byte
    Dim i As Integer
    
    tb_output.Text = ""
    dataLen = CByte(packetLength)
    
    Sleep 500
    stsResult = PCANBasic.CAN_Read(m_PcanHandle, CANMsg, CANTimeStamp)
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        While CANMsg.LEN <> dataLen
            stsResult = PCANBasic.CAN_Read(m_PcanHandle, CANMsg, CANTimeStamp)
        Wend
        For i = 0 To CANMsg.LEN - 1
            If setNetIdCalled = True Then
                tb_output.Text = tb_output.Text & Format(Hex(CANMsg.DATA(i)), "00")
            Else
                tb_output.Text = tb_output.Text & Chr(CANMsg.DATA(i))
                If Chr(CANMsg.DATA(i)) = "0" Then
                    tb_output.Text = tb_output.Text & "0"
                End If
            End If
        Next i
        If readSNCalled = True And setNetIdCalled = False Then
            serialNumber = Trim(Mid(tb_output.Text, 2, 7))
            Call setNetID
        End If
    Else
        tb_output.Text = "Read Message Failed"
    End If
End Sub

Private Sub setNetID()
    s_netID = CLng("&H" & "400000") Or CLng(serialNumber)
    setNetIdCalled = True
End Sub

Private Sub dacZero()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd As String
    Dim msg() As String
    Dim msg1() As String
    Dim dataLen As Integer
    
    cmnd = "d1"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_zero.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 5, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub dacSpan()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd As String
    Dim msg() As String
    Dim msg1() As String
    Dim dataLen As Integer
    
    cmnd = "d2"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_span.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 5, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub dacGain()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd As String
    Dim msg() As String
    Dim msg1() As String
    Dim dataLen As Integer
    
    cmnd = "d3"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_gain.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 5, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub temp_btn_Click()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim cmnd As String
    Dim msg() As String
    Dim msg1() As String
    Dim dataLen As Integer
    
    cmnd = "t"
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    msg1 = str2Array(cmnd)
    ReDim msg(UBound(msg1))
    For i = 0 To UBound(msg1)
        msg(i) = string2Hex(msg1(i))
    Next i
    
    dataLen = UBound(msg1) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Chr(CANMsg.DATA(i))
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_temp.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 3, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub writeZero()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim packet As String
    Dim msg() As String
    Dim dataLen As Integer
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    If CLng(tb_zero.Text) >= 0 Or CLng(tb_zero.Text) <= 65536 Then
        packet = Hex(Asc("D")) & Hex(Asc("1")) & littleEndian(Format(Hex(CLng(tb_zero.Text)), "0000"))
    Else
        tb_zero.Text = "Invalid Val"
        Exit Sub
    End If
    
    msg = str2ArrayPair(packet)

    dataLen = UBound(msg) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Format(Hex(CANMsg.DATA(i)), "00")
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_zero.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 5, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub writeSpan()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim packet As String
    Dim msg() As String
    Dim dataLen As Integer
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    If CLng(tb_span.Text) >= 0 Or CLng(tb_span.Text) <= 65536 Then
        packet = Hex(Asc("D")) & Hex(Asc("2")) & littleEndian(Format(Hex(CLng(tb_span.Text)), "0000"))
    Else
        tb_span.Text = "Invalid Val"
        Exit Sub
    End If
    
    msg = str2ArrayPair(packet)

    dataLen = UBound(msg) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Format(Hex(CANMsg.DATA(i)), "00")
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_span.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 5, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

Private Sub writeGain()
    Dim CANMsg As TPCANMsg
    Dim stsResult As Long
    Dim packet As String
    Dim msg() As String
    Dim dataLen As Integer
    
    tb_input.Text = ""
    tb_output.Text = ""
    tb_messageSent.Text = ""
    tb_messageSent.BackColor = vbWhite
    
    If CLng(tb_gain.Text) >= 0 Or CLng(tb_gain.Text) <= 65536 Then
        packet = Hex(Asc("D")) & Hex(Asc("3")) & littleEndian(Format(Hex(CLng(tb_gain.Text)), "0000"))
    Else
        tb_gain.Text = "Invalid Val"
        Exit Sub
    End If
    
    msg = str2ArrayPair(packet)

    dataLen = UBound(msg) + 1
    
    CANMsg.ID = s_netID
    
    CANMsg.LEN = CByte(dataLen)
    
    CANMsg.MsgType = PCANBasic.PCAN_MESSAGE_EXTENDED
    
    For i = 0 To UBound(msg)
        If msg(i) = "" Then
            msg(i) = "00"
        End If
        CANMsg.DATA(i) = CByte("&H" & msg(i))
    Next i
    
    If connected = True And mm_enabled = True And setNetIdCalled = True Then
        stsResult = PCANBasic.CAN_Write(m_PcanHandle, CANMsg)
    Else
        stsResult = "999999"
    End If
    
    For i = 0 To UBound(CANMsg.DATA)
        tb_input.Text = tb_input.Text & " " & Format(Hex(CANMsg.DATA(i)), "00")
    Next i
    
    If stsResult = PCANBasic.PCAN_ERROR_OK Then
        tb_messageSent.Text = "Message Sent Successfully"
        tb_messageSent.BackColor = vbGreen
        If connected = True And mm_enabled = True And setNetIdCalled = True Then
            Call readMessages(8)
            tb_gain.Text = CLng("&H" & littleEndian(Trim(Mid(tb_output.Text, 5, 4))))
        End If
    Else
        tb_messageSent.Text = "Message Sent Failed"
        tb_messageSent.BackColor = vbRed
        tb_output.Text = stsResult
    End If
End Sub

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

Private Function str2Array(xString As String) As String()
Dim tempArray() As String
ReDim tempArray(Len(xString) - 1)

For i = 1 To Len(xString)
    tempArray(i - 1) = Mid(xString, i, 1)
Next i

str2Array = tempArray
End Function

Private Function string2Hex(xString As String) As String
For i = 1 To Len(xString)
    string2Hex = string2Hex & Hex$(Asc(Mid(xString, i, 1))) & Space$(1)
    On Error Resume Next
    DoEvents
    Next i
    string2Hex = Mid$(string2Hex, 1, Len(string2Hex) - 1)
End Function
