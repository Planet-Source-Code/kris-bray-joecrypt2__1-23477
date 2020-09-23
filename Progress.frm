VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Initializing Sequence"
   ClientHeight    =   2925
   ClientLeft      =   3270
   ClientTop       =   3810
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7935
      _cx             =   13996
      _cy             =   3836
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ExPro = True
Form1.Label3.Caption = "Status: Process Canceled"

Form2.Hide

End Sub

Private Sub Form_Load()
Flash1.Movie = App.Path & "\binaryrandom.swf"
Flash1.Play

Lenoftxt1 = Len(MDIForm1.ActiveForm.Text1.Text)
ProgressBar1.Max = Lenoftxt1

ProgressBar1.Value = 0
End Sub

Private Sub Timer1_Timer()
If Image1.Visible = False Then Image1.Visible = True: Image2.Visible = False: Exit Sub
If Image2.Visible = False Then Image2.Visible = True: Image1.Visible = False


End Sub
