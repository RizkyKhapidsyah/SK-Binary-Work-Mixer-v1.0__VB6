VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "BinaryWork Mixer  OCX samples"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Advanced project"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple project"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show

Unload Form3
Set Form3 = Nothing

End Sub

Private Sub Command2_Click()

Form1.Show

Unload Form3
Set Form3 = Nothing

End Sub
