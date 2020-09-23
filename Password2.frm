VERSION 5.00
Begin VB.Form form4 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter Password For Decryption"
   ClientHeight    =   660
   ClientLeft      =   3450
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
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub com1_Click()
password2 = Text1.Text
If password2 = password Then
PassEnt = False
Call FinishDec

form4.Hide
    Else
          
        
        MsgBox "Incorrect Password! Please Try Again", , "Password Error"
        
End If

End Sub

Private Sub com2_Click()
PassEnt = False
Form2.Hide
form4.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call com1_Click
End Sub
