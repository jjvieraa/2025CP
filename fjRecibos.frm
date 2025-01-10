VERSION 5.00
Begin VB.Form fjRecibos 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Recibos"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   Icon            =   "fjRecibos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmGenerar 
      BackColor       =   &H0080C0FF&
      Caption         =   "Generar"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "No. Cobrador:"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "No Recibo:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   1335
   End
End
Attribute VB_Name = "fjRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cREC As New clsRecibos








'=================================
Private Sub Form_Load()
'=================================
txt(0).Text = cREC.mfTomaNumRecibo
End Sub



'=================================
Private Sub cmGenerar_Click()
'=================================
If Not IsNumeric(txt(1).Text) Then
    txt(1).SetFocus
    Exit Sub
End If

End Sub



'======================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
'======================================================
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"        ' COMO SI PULSARA ENTER
    End If
End Sub

'=====================================================================
Private Sub txt_GotFocus(Index As Integer)
'=====================================================================
   
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index).Text)
End Sub



