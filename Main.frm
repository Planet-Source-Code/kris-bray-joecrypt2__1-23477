VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Untitled"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   10110
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   5610
      Width           =   9885
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   4905
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   270
      Width           =   9870
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7305
      TabIndex        =   4
      Top             =   5310
      Width           =   2670
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Message Text"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   30
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   5340
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocChange As Boolean

Sub DocCF()
DocChange = False
End Sub



Private Sub Form_GotFocus()
MDIForm1.Caption = "JoeCrypt - " & Me.Caption
End Sub

Private Sub Form_Load()
DocChange = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If DocChange = True Then
        confclose = MsgBox("One or more documents have changed! Do you wish to lose the changes? ", vbYesNo, "Confirm Exit")
    If confclose = vbYes Then
        Unload Me
            Else
        CanUnload = True
        Cancel = 1
    End If
End If
End Sub

Private Sub Text1_Change()
DocChange = True
End Sub
