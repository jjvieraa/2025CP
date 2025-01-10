VERSION 5.00
Begin VB.Form fjAviso2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Aviso"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Continuar"
      Height          =   435
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "fjAviso2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

