VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThisIsMyForm!"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "F2"
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Dim f As Form2
    Set f = New Form2
    Load f
    f.Show vbModal
End Sub

Private Sub Form_Load()
    MsgBox "This is Load, caption was " & Me.Caption
    Me.Caption = "LOL!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MsgBox "QueryUnload"
End Sub

Private Sub Command1_Click()
    MsgBox "Command1 clicked!"
    Label1.Caption = "concha la madre " & Now
End Sub
