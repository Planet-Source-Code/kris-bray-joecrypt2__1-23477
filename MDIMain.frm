VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "JoeCrypt 2.0"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8370
      TabIndex        =   0
      Top             =   0
      Width           =   8430
      Begin VB.CommandButton Command6 
         Caption         =   "Chat"
         Height          =   570
         Left            =   2310
         Picture         =   "MDIMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Encrypted Chat Over The Internet"
         Top             =   15
         Width           =   465
      End
      Begin VB.CommandButton Command5 
         Height          =   570
         Left            =   1845
         Picture         =   "MDIMain.frx":07A6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Send Current Document through E-mail"
         Top             =   15
         Width           =   465
      End
      Begin VB.CommandButton Command4 
         Height          =   570
         Left            =   1380
         Picture         =   "MDIMain.frx":0F4C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Generate Key Now"
         Top             =   15
         Width           =   465
      End
      Begin VB.CommandButton Command3 
         Height          =   570
         Left            =   915
         Picture         =   "MDIMain.frx":16F2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Options"
         Top             =   15
         Width           =   465
      End
      Begin VB.CommandButton Command2 
         Height          =   570
         Left            =   450
         Picture         =   "MDIMain.frx":1E98
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Decrypt Current Document"
         Top             =   15
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         Height          =   570
         Left            =   -15
         Picture         =   "MDIMain.frx":2D9E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Encrypt Current Document"
         Top             =   15
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Note: It is recommended that you save your documents in JoeCrypt Format (*.JCD) especially after encrypting a document."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   510
         Left            =   2760
         TabIndex        =   7
         Top             =   60
         Width           =   5610
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   150
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu cmdNew 
         Caption         =   "New"
      End
      Begin VB.Menu cmdopen 
         Caption         =   "Open"
      End
      Begin VB.Menu cmdclosedoc 
         Caption         =   "Close"
      End
      Begin VB.Menu cmdSave 
         Caption         =   "Save"
      End
      Begin VB.Menu cmdopenkey 
         Caption         =   "Load Key"
      End
      Begin VB.Menu cmdSaveKey 
         Caption         =   "Save Key"
      End
      Begin VB.Menu CmdExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu cmdClearDoc 
         Caption         =   "Clear Document"
      End
      Begin VB.Menu cmdEncrypt 
         Caption         =   "Encrypt"
      End
      Begin VB.Menu cmdDecrypt 
         Caption         =   "Decrypt"
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "Options"
      Begin VB.Menu subEncMeth 
         Caption         =   "Encryption Method"
         Begin VB.Menu cmdrndchr 
            Caption         =   "Random Characters"
         End
         Begin VB.Menu cmd3d 
            Caption         =   "3 Digit Numbers"
         End
      End
      Begin VB.Menu SubSound 
         Caption         =   "Sound"
         Begin VB.Menu chkSoundOn 
            Caption         =   "Sound On"
            Checked         =   -1  'True
         End
         Begin VB.Menu chkSoundOff 
            Caption         =   "Sound Off"
         End
      End
      Begin VB.Menu cmdgenerate 
         Caption         =   "Generate Key Now"
      End
      Begin VB.Menu genkeyonenc 
         Caption         =   "Generate Key On Encrypt"
         Checked         =   -1  'True
      End
      Begin VB.Menu passifcmd 
         Caption         =   "Password Documents"
         Checked         =   -1  'True
      End
      Begin VB.Menu cmdsndEmail 
         Caption         =   "Send E-Mail"
      End
      Begin VB.Menu cmdEChat 
         Caption         =   "Encrypted Chat"
      End
      Begin VB.Menu cmdkeyindoc 
         Caption         =   "Include Key In Document"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MnuWindowList 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu cmdCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu CmdTile 
         Caption         =   "&Tile"
      End
      Begin VB.Menu cmdArrange 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "Help"
      Begin VB.Menu cmdAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkSoundOff_Click()
chkSoundOff.Checked = True
chkSoundOn.Checked = False
SOUND_ON = False

End Sub

Private Sub cmdArrange_Click()
MDIForm1.Arrange vbArrangeIcons

End Sub

Private Sub cmdCascade_Click()
MDIForm1.Arrange vbCascade

End Sub

Private Sub cmdclosedoc_Click()
On Error Resume Next
Unload MDIForm1.ActiveForm
End Sub

Private Sub cmdDecrypt_Click()
DecryptMsg
End Sub

Private Sub cmdEncrypt_Click()
EncryptMsg
End Sub

Private Sub cmdkeyindoc_Click()
If cmdkeyindoc.Checked = True Then
    cmdkeyindoc.Checked = False
    KeyInDoc = False
    cmdSaveKey.Enabled = True
                Else
    cmdkeyindoc.Checked = True
    cmdSaveKey.Enabled = False
    KeyInDoc = True
End If

End Sub

Private Sub cmdNew_Click()
    ' Create a new instance of Form1, called NewDoc.
    Dim NewDoc As New Form1
    ' Display the new form.
    NewDoc.Show
NewDoc.Height = 6390
NewDoc.Width = 10230
End Sub

Private Sub CmdTile_Click()
MDIForm1.Arrange vbTileHorizontal

End Sub

Private Sub Command1_Click()
On Error GoTo encerror
    If MDIForm1.ActiveForm.Text1.Text <> "" Then EncryptMsg
encerror:
    If Err = 91 Then MsgBox "No Document is currently Open and Active", , "No Document to Encrypt"
    
End Sub

Private Sub Command2_Click()
On Error GoTo encerror
    If MDIForm1.ActiveForm.Text1.Text <> "" Then DecryptMsg
encerror:
    If Err = 91 Then MsgBox "No Document is currently Open and Active", , "No Document to Decrypt"
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub

Private Sub Command4_Click()
On Error GoTo encerror
    If MDIForm1.ActiveForm.Text1.Text <> "" Then KeyGen
encerror:
    If Err = 91 Then MsgBox "No Document is currently Open and Active", , "No Document to generate a key for"
    
End Sub

Private Sub Command5_Click()
On Error GoTo encerror
    FrmEmail.Show
    FrmEmail.txtEmailBodyOfMessage.Text = MDIForm1.ActiveForm.Text1.Text

encerror:
    If Err = 91 Then MsgBox "No Document is currently Open and Active", , "No Document to send in e-mail": FrmEmail.Hide
    
End Sub

Private Sub Command6_Click()
Chat1.Show
End Sub

Private Sub MDIForm_Load()
If Command <> "" Then
    CD1.filename = Command
    Call cmdopen_Click
End If
SOUND_ON = True
ChrEvent = True
GenK = True
PassIf = 0
PassDoc = True
KeyInDoc = True
CanUnload = False
End Sub


Private Sub chkSoundOn_Click()
chkSoundOn.Checked = True
chkSoundOff.Checked = False
SOUND_ON = True

End Sub

Private Sub cmd3d_Click()
ChrEvent = False
cmdrndchr.Checked = False
cmd3d.Checked = True

End Sub

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdcleardoc_Click()
MDIForm1.ActiveForm.Text1.Text = ""
PassIf = 0
password = ""
End Sub

Private Sub cmdEChat_Click()
Chat1.Show
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdgenerate_Click()
Call KeyGen
End Sub

Private Sub cmdopen_Click()
Dim IsKeyInDoc As Boolean
If Command <> "" Then GoTo 5
CD1.Filter = "JoeCrypt Documents (*.jcd)|*.jcd|All Files (*.*)|*.*"
CD1.ShowOpen
5
 If CD1.filename <> "" Then
        If Right(CD1.filename, 3) <> "jcd" Then
            Open CD1.filename For Input As #1
                Do Until EOF(1)
                    Line Input #1, lineoftext
                    alltext = alltext & lineoftext & Chr(13) + Chr(10)
                Loop
            Close #1
                Else
    Open CD1.filename For Input As #1
        Input #1, IsKeyInDoc
            If IsKeyInDoc = True Then Line Input #1, keyfil
    
    Do Until EOF(1)
    Line Input #1, lineoftext
        alltext = alltext & lineoftext & Chr(13) + Chr(10)
    Loop
    Close #1
        End If
        ' Create a new instance of Form1, called NewDoc.
         Dim NewDoc As New Form1
          ' Display the new form.
                NewDoc.Show
                    NewDoc.Height = 6390
                    NewDoc.Width = 10230
        MDIForm1.ActiveForm.Caption = "JoeCrypt - " & CD1.filename
    MDIForm1.ActiveForm.Text1.Text = alltext
    If IsKeyInDoc = True Then MDIForm1.ActiveForm.Text2.Text = keyfil
 End If
End Sub

Private Sub cmdopenkey_Click()
CD1.Filter = "JoeCrypt Key (*.jck)|*.jck"
CD1.ShowOpen
    
        If CD1.filename <> "" Then
            Open CD1.filename For Input As #1
            Do Until EOF(1)
            Line Input #1, lineot
            allt = allt & lineot
            Loop
            Close #1
            MDIForm1.ActiveForm.Text2.Text = allt
        End If
    
End Sub

Private Sub cmdrndchr_Click()
ChrEvent = True
cmdrndchr.Checked = True
cmd3d.Checked = False

End Sub

Private Sub cmdsave_Click()
CD1.Filter = "JoeCrypt Documents (*.JCD)|*.jcd|All Files (*.*)|*.*"
CD1.ShowSave
    If CD1.filename <> "" Then
        If Right(CD1.filename, 3) <> "jcd" Then
              MDIForm1.ActiveForm.DocCF
        Open CD1.filename For Output As #1
            Print #1, MDIForm1.ActiveForm.Text1.Text
        Close #1
            MDIForm1.ActiveForm.Caption = CD1.filename
                CD1.filename = ""
                
            Else
       
       MDIForm1.ActiveForm.DocCF
        Open CD1.filename For Output As #1
            Write #1, KeyInDoc
            If KeyInDoc = True Then Print #1, MDIForm1.ActiveForm.Text2.Text
            Print #1, MDIForm1.ActiveForm.Text1.Text
        Close #1
            MDIForm1.ActiveForm.Caption = CD1.filename
                CD1.filename = ""
        End If
    End If


End Sub

Private Sub cmdsavekey_Click()
CD1.Filter = "JoeCrypt Key (*.jck)|*.jck"
CD1.ShowSave
    
    If CD1.filename <> "" Then
        Open CD1.filename For Output As #1
            Print #1, MDIForm1.ActiveForm.Text2.Text
        Close #1
    End If

End Sub





Private Sub cmdSndEmail_Click()
FrmEmail.Show
FrmEmail.txtEmailBodyOfMessage.Text = MDIForm1.ActiveForm.Text1.Text

End Sub

Private Sub Decrypt_Click()
DecryptMsg
End Sub

Private Sub Encrypt_Click()
EncryptMsg
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub genkeyonenc_Click()
If genkeyonenc.Checked = True Then
    genkeyonenc.Checked = False
    GenK = False
        Else
    genkeyonenc.Checked = True
    GenK = True

End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
Image2.Visible = True
Call Encrypt_Click
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image1.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image1.Visible = True
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image4.Visible = True
Call Decrypt_Click
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image3.Visible = True
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image3.Visible = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If Cancel <> 1 Then End
End Sub

Private Sub MnuFile_Click()
If KeyInDoc = True Then cmdSaveKey.Enabled = False

End Sub

Private Sub passifcmd_Click()
If passifcmd.Checked = True Then
    passifcmd.Checked = False
    PassDoc = False
        PassIf = 0
        Else
    passifcmd.Checked = True
    PassDoc = True
    PassIf = 1
End If

    

End Sub


