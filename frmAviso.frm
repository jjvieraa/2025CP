VERSION 5.00
Begin VB.Form fjAviso 
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmAviso.frx":0000
   ScaleHeight     =   870
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   255
   End
   Begin VB.Label mMensj 
      Alignment       =   2  'Center
      Caption         =   "Probando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   270
      TabIndex        =   0
      Top             =   75
      Width           =   2820
   End
End
Attribute VB_Name = "fjAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ni

Private Sub Form_Load()
ni = 0
End Sub

Private Sub Timer1_Timer()
ni = ni + 1
If ni = 3 Then Unload Me
End Sub
