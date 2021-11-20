VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Winsock Control Example"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      ToolTipText     =   "Send Information"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtSend 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      MaxLength       =   1024
      TabIndex        =   5
      ToolTipText     =   "Information to Send"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2880
      Top             =   960
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "Connect to a remote computer"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Listen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Listen for incoming connections"
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Remote Information:"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   3615
      Begin VB.TextBox txtIP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Remote Computer's IP Address"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label SendStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label lab 
         BackStyle       =   0  'Transparent
         Caption         =   "Winsock Connecting Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "The Connecting Control's Status"
         Top             =   675
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Port: 554"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         ToolTipText     =   "Port: 554 -> cuz I wanted it that way!"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lab 
         Caption         =   "Remote IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local Information:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.Label ListenStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label lab 
         BackStyle       =   0  'Transparent
         Caption         =   "Winsock Listening Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "The Listening Control's Status"
         Top             =   675
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Port: 554"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "Port: 554 -> cuz I wanted it that way!"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label labLocalIP 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         ToolTipText     =   "Your computer's IP Address... click to Copy to Clipboard"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lab 
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   3720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   554
   End
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   3720
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   554
   End
   Begin VB.Label labMsg 
      Alignment       =   2  'Center
      Caption         =   "[ no messages to display ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Last recieved message"
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label lab 
      Caption         =   "Send:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
If cmdConnect.Caption = "&Connect" Then
    If Len(txtIP.Text) = 0 Then
        Beep
        txtIP.SetFocus
        Exit Sub
        End If
    If txtIP.Text = sckListen.LocalIP Then
        Beep
        MsgBox "You cannot connect to yourself.", vbInformation
        Exit Sub
        End If
    cmdConnect.Caption = "&Close"
    cmdListen.Enabled = False
    txtSend.Enabled = True
    cmdSend.Enabled = True
    sckSend.Connect txtIP.Text, 554
   Else
    sckSend.Close
    cmdListen.Enabled = True
    cmdConnect.Caption = "&Connect"
    txtSend.Enabled = False
    cmdSend.Enabled = False
    End If
End Sub

Private Sub cmdListen_Click()
If cmdListen.Caption = "&Listen" Then
    cmdListen.Caption = "&Close"
    cmdConnect.Enabled = False
    txtSend.Enabled = True
    cmdSend.Enabled = True
    sckListen.Listen
   Else
    sckListen.Close
    cmdListen.Caption = "&Listen"
    cmdConnect.Enabled = True
    txtSend.Enabled = False
    cmdSend.Enabled = False
    End If
End Sub


Private Sub cmdSend_Click()
If cmdListen.Enabled = True Then
    sckListen.SendData (txtSend.Text)
   Else
    sckSend.SendData CStr(txtSend.Text)
    End If
txtSend.Text = ""
End Sub

Private Sub Form_Load()
labLocalIP.Caption = sckListen.LocalIP
End Sub

Private Sub labLocalIP_Click()
Clipboard.SetText labLocalIP.Caption
Beep
End Sub

Private Sub sckListen_Close()
txtIP.Locked = False
cmdListen.Caption = "&Close"
Call cmdListen_Click
End Sub

Private Sub sckListen_Connect()
txtIP.Text = sckListen.RemoteHostIP
txtIP.Locked = True

End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
If sckListen.State = 2 Then sckListen.Close
sckListen.Accept (requestID)
txtSend.Enabled = True
cmdSend.Enabled = True
End Sub


Private Sub sckListen_DataArrival(ByVal bytesTotal As Long)
sckListen.GetData buffer$, vbString
Beep
labMsg.Caption = buffer$
End Sub


Private Sub sckSend_Close()
cmdSend.Caption = "&Cancel"
Call cmdSend_Click
End Sub

Private Sub sckSend_Connect()
cmdSend.Enabled = True
txtSend.Enabled = True
End Sub


Private Sub sckSend_DataArrival(ByVal bytesTotal As Long)
sckSend.GetData buffer$, vbString
Beep
labMsg.Caption = buffer$
End Sub


Private Sub Timer1_Timer()
If sckListen.State = 0 Then ListenStatus.Caption = "Closed"
If sckListen.State = 1 Then ListenStatus.Caption = "Open"
If sckListen.State = 2 Then ListenStatus.Caption = "Listening"
If sckListen.State = 3 Then ListenStatus.Caption = "Connection Pending"
If sckListen.State = 4 Then ListenStatus.Caption = "Resolving Host"
If sckListen.State = 5 Then ListenStatus.Caption = "Host Resolved"
If sckListen.State = 6 Then ListenStatus.Caption = "Connecting"
If sckListen.State = 7 Then ListenStatus.Caption = "Connected"
If sckListen.State = 8 Then ListenStatus.Caption = "No Carrier"
If sckListen.State = 9 Then ListenStatus.Caption = "Error"
If sckSend.State = 0 Then SendStatus.Caption = "Closed"
If sckSend.State = 1 Then SendStatus.Caption = "Open"
If sckSend.State = 2 Then SendStatus.Caption = "Listening"
If sckSend.State = 3 Then SendStatus.Caption = "Connection Pending"
If sckSend.State = 4 Then SendStatus.Caption = "Resolving Host"
If sckSend.State = 5 Then SendStatus.Caption = "Host Resolved"
If sckSend.State = 6 Then SendStatus.Caption = "Connecting"
If sckSend.State = 7 Then SendStatus.Caption = "Connected"
If sckSend.State = 8 Then SendStatus.Caption = "No Carrier"
If sckSend.State = 9 Then SendStatus.Caption = "Error"
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdConnect_Click
End Sub


Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSend_Click
End Sub


