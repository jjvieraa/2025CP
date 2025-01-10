VERSION 5.00
Begin VB.Form fjAcerca 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de CP"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5940
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5577.967
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      ClipControls    =   0   'False
      Height          =   1995
      Left            =   240
      Picture         =   "fjAcerca.frx":0000
      ScaleHeight     =   1359.015
      ScaleMode       =   0  'User
      ScaleWidth      =   1422.225
      TabIndex        =   1
      Top             =   240
      Width           =   2085
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2625
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Info. del sistema..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   4320
      MaskColor       =   &H80000013&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3075
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Última Actualiz.:"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nov 2010"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Program: Juan Viera     Tel. 46222626       jviera@adinet.com.uy"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Círculo Policial. Rivera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   3405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Versión 2.91"
      Height          =   225
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1245
   End
End
Attribute VB_Name = "fjAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ultima modif: 11/10/02

Option Explicit

Private Sub cmdOK_Click()
  Unload Me

End Sub

Private Sub Form_Load()
    'Me.Caption = "Acerca de " & App.Title
    'lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
End Sub

