VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Passwords"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
      Begin VB.CheckBox Check1 
         Caption         =   "Use Passwords"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sounds"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
      Begin VB.OptionButton Option4 
         Caption         =   "Sound Off"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Sound On"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encryption"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "3 Digit Numbers"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Random Characters"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If MDIForm1.passifcmd.Checked = True Then
        Check1.Value = vbUnchecked
    MDIForm1.passifcmd.Checked = False
    PassDoc = False
        PassIf = 0
        Else
        Check1.Value = vbChecked
    MDIForm1.passifcmd.Checked = True
    PassDoc = True
    PassIf = 1
End If

End Sub

Private Sub Option1_Click()
ChrEvent = True
MDIForm1.cmdrndchr.Checked = True
MDIForm1.cmd3d.Checked = False

End Sub

Private Sub Option2_Click()
ChrEvent = False
MDIForm1.cmdrndchr.Checked = False
MDIForm1.cmd3d.Checked = True

End Sub

Private Sub Option3_Click()
MDIForm1.chkSoundOn.Checked = True
MDIForm1.chkSoundOff.Checked = False
SOUND_ON = True

End Sub

Private Sub Option4_Click()
MDIForm1.chkSoundOff.Checked = True
MDIForm1.chkSoundOn.Checked = False
SOUND_ON = False

End Sub
