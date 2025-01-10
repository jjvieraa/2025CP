VERSION 5.00
Begin VB.Form fjPideTCbio 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Tasa Cambio"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,#0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,#0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1755
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,#0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1185
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,#0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   615
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Aceptar"
      Height          =   270
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Cancelar"
      Height          =   270
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3030
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Unidad R."
      Height          =   165
      Index           =   3
      Left            =   195
      TabIndex        =   9
      Top             =   2385
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "P.Argentino"
      Height          =   165
      Index           =   2
      Left            =   195
      TabIndex        =   7
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Real"
      Height          =   165
      Index           =   1
      Left            =   195
      TabIndex        =   6
      Top             =   1290
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Dólar"
      Height          =   165
      Index           =   0
      Left            =   195
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "fjPideTCbio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=====================================================================
Private Sub Command1_Click()
'=====================================================================
Unload Me
End Sub


'=====================================================================
Private Sub Command2_Click()
'=====================================================================
vpTCD = Text1(0).Text
vpTCR = Text1(1).Text
vpTCA = Text1(2).Text
vpTCU = Text1(3).Text
vpbVieneDeMuestraTabla = True
Me.Hide
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Me.Refresh
DoEvents
End Sub

'=====================================================================
Private Sub Text1_GotFocus(Index As Integer)
'=====================================================================
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub



'=====================================================================
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'=====================================================================
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

'=====================================================================
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
'=====================================================================
Text1(Index).Text = Format(Text1(Index).Text, "#,#0.00")
End Sub

