VERSION 5.00
Begin VB.Form frmPrueba 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3345
      TabIndex        =   0
      Top             =   1350
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tc As New clsTCambio


Private Sub Command1_Click()
tc.vdFecha = CDate("12/11/2002")
tc.mfBuscaTCambio (1)
End Sub
