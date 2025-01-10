VERSION 5.00
Begin VB.Form fjPideDato 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sección"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0FF&
   ForeColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Canc"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
      Width           =   2295
   End
End
Attribute VB_Name = "fjPideDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
    vpbCancel = True
    vpsPideDato = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    vpbCancel = False
    vpsPideDato = Text1.Text
    Unload Me
End Sub



Private Sub Text1_Validate(Cancel As Boolean)
If IsNull(Text1.Text) Or Len(Trim(Text1.Text)) = 0 Then
    Cancel = True
End If
End Sub
