VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Chat1 
   BackColor       =   &H00000000&
   Caption         =   "Encrypted Chat"
   ClientHeight    =   5310
   ClientLeft      =   1500
   ClientTop       =   525
   ClientWidth     =   8505
   Icon            =   "EChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8505
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6255
      Top             =   930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Client"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2385
      TabIndex        =   16
      Top             =   1710
      Width           =   780
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1050
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Status"
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1470
      Width           =   2160
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Idle.."
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   75
         TabIndex        =   13
         Top             =   210
         Width           =   1950
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2445
      TabIndex        =   11
      Top             =   1050
      Width           =   780
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3420
      TabIndex        =   9
      Top             =   225
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7395
      TabIndex        =   8
      Top             =   4920
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2010
      Width           =   8295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   105
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4935
      Width           =   7125
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1035
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get"
      Height          =   375
      Left            =   2460
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Host Name:"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   3420
      TabIndex        =   14
      Top             =   795
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Name:"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   3435
      TabIndex        =   10
      Top             =   15
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Outgoing..."
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   4605
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Destination IP address"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   825
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Local IP"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   135
      TabIndex        =   2
      Top             =   45
      Width           =   945
   End
End
Attribute VB_Name = "Chat1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DataSnd, DataRec
Public RemoteName, UserName

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    Winsock1.Close
    Command3.Enabled = True
        Else
    Command2.Enabled = False
    Winsock1.Close
    Winsock1.Listen
End If
End Sub

Private Sub Command1_Click()
Text1.Text = Winsock1.LocalIP
End Sub

Private Sub Command2_Click()


' Send the data that is keyed into this textbox
' to the other user.


If Text3.Text <> "" Then
       If Text4.Text <> "" Then Text4.Text = Text4.Text & vbNewLine & "<" & UserName & "> " & Text3.Text
       If Text4.Text = "" Then Text4.Text = Text4.Text & "<" & UserName & "> " & Text3.Text

    ChatEnc
    Winsock1.SendData "909" & DataSnd
    DataSnd = ""
    Text3.Text = ""
End If

End Sub

Private Sub Command3_Click()
If Command3.Caption = "Connect" Then
    If Text5.Text <> "" And Text2.Text <> "" Then
        Winsock1.RemoteHost = Text2.Text
        Winsock1.RemotePort = 2002
        Winsock1.Connect
    
    
            Else
        If Text5.Text = "" Then MsgBox "Please Enter Your Name in the 'Name' Field", vbOKOnly, "Connection Error" Else
        If Text2.Text = "" Then MsgBox "Please enter a valid IP address", , "Connection Error"
End If
        Else
    Winsock1.Close
    Text3.Enabled = False
    Text4.Enabled = False
    Command2.Enabled = False
    Command3.Caption = "Connect"
End If
End Sub

Private Sub Form_Load()

Winsock1.Protocol = sckTCPProtocol
Winsock1.LocalPort = 2002
Text1.Text = Winsock1.LocalIP
If Check1.Value = vbUnchecked Then
    Winsock1.Listen
    Command3.Enabled = False
End If

End Sub





Private Sub Text2_Change()
' Set the remote host equivalent to the
' IP address in Text3


End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

Text3 = Text3

'if Enter Key is pressed Encrypt and send
If KeyAscii = 13 Then
       
       If Text4.Text <> "" Then Text4.Text = Text4.Text & vbNewLine & "<" & UserName & "> " & Text3.Text
       If Text4.Text = "" Then Text4.Text = Text4.Text & "<" & UserName & "> " & Text3.Text

    Label5.Caption = "Encrypting..."
    ChatEnc
    Label5.Caption = "Sending..."
    Winsock1.SendData "909" & DataSnd
    DataSnd = ""
    Text3.Text = ""
    Exit Sub
End If
DataSnd = DataSnd & Chr(KeyAscii)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

' Set the textbox equivalent to itself.

Text2 = Text2

' Get the data from the other user as he/she is typing.

Winsock1.GetData (KeyAscii)

End Sub

Private Sub Text5_Change()
UserName = Text5.Text
End Sub

Private Sub Winsock1_Close()
Label5.Caption = "Disconnected."
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData ("101" & UserName)
Label5.Caption = "Connected"
Command3.Caption = "Disconnect"
Text3.Enabled = True
Text4.Enabled = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    'let User know someone is trying to connect"
    Label5.Caption = "Connection Requested"
    ' Check if the control's State is closed. If not,
    ' close the connection before accepting the new
    ' connection.
    If Winsock1.State <> sckClosed Then Winsock1.Close
    ' Accept the request with the requestID
    ' parameter.
    Winsock1.Accept requestID
    'Set Destination IP to Remote IP

    Text2.Text = Winsock1.RemoteHost
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

' Make a string that will represent the data
' that is being sent from one computer to
' yours.

Dim strData As String
strData = ""
' Then tell winsock to get that data and import it.

Winsock1.GetData strData, vbString

' Changes the remote IP to where the incoming data
' came from.

Text2.Text = Winsock1.RemoteHostIP



If Mid(strData, 1, 3) = "909" Then
    DataRec = Right(strData, Len(strData) - 3)
    Label5.Caption = "Decrypting"
    DecChat
   If Text4.Text <> "" Then Text4.Text = Text4.Text & vbNewLine & "<" & RemoteName & "> " & DataRec
   If Text4.Text = "" Then Text4.Text = Text4.Text & "<" & RemoteName & "> " & DataRec
End If

If Mid(strData, 1, 3) = "101" Then
    RemoteName = Right(strData, Len(strData) - 3)
    Text6.Text = RemoteName
    
         If UserName = "" Then UserName = InputBox("Please Enter a Name for yourself", "Enter Nick"): Text5.Text = UserName
        
    Winsock1.SendData ("102" & UserName)
End If

If Mid(strData, 1, 3) = "102" Then
    RemoteName = Right(strData, Len(strData) - 3)
    Text6.Text = RemoteName
    
    
    Command2.Enabled = True
    Winsock1.SendData ("103")
End If

If Mid(strData, 1, 3) = "103" Then
    Label5.Caption = "Connected"
    Text3.Enabled = True
    Text4.Enabled = True
    Command2.Enabled = True
End If

If Asc(strData) = 8 And Len(Text4.Text) > 0 Then
   Text4.Text = Mid(Text4.Text, 1, (Len(Text4.Text) - 1))
Else
   
End If

' If the user presses the enter key, then it
' will register on both ends and send the data.

If Asc(strData) = 13 Then
   Text4.Text = Text4.Text & vbNewLine
End If

' This will make the textbox the other person is
' talking through automatically move down when it
' gets to the bottom of the textbox.

Text4.SelStart = Len(Text4.Text)

End Sub
