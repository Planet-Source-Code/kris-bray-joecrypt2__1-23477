VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter Password For Encryption"
   ClientHeight    =   660
   ClientLeft      =   3270
   ClientTop       =   4170
   ClientWidth     =   4860
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton com2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton com1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   105
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub com1_Click()
Dim Chngfrmtst As String
password = Text1.Text
PassLen = Len(password)
Chngfrmtst = PassLen
PassLenoLen = Len(Chngfrmtst)
PassIf = 1
PassEnt = True
Form3.Hide
EncryptMsg

End Sub

Private Sub com2_Click()
MsgBox "If you do NOT wish to password this document you must UNcheck Password Documents in the Options Menu!", , "JoeCrypt Info"

Form3.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call com1_Click

End Sub
