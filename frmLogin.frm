VERSION 5.00
Begin VB.Form frmjLogin 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada"
   ClientHeight    =   1545
   ClientLeft      =   4380
   ClientTop       =   4575
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "vr 2.88"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmjLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'NIVELES
'6: ADMINISTRADOR
'5: USUARIO TOTAL EXCEPTO FJDATOS
'1: SOLO INGRESA SOCIOS


Option Explicit
Dim cuenta As Byte


'======================================================
Private Sub txtUserName_Validate(Cancel As Boolean)
 '======================================================
   Dim sCriterio As String
    
    If Not IsNumeric(txtUserName.Text) Then
        Cancel = True
    Else
        'si no la encuentra
        If Not TomaDatosFUncionario(txtUserName.Text) Then
            MsgBox ("Error en el código")
            CierraLaBase
            End
        End If
        Label1.Caption = vptNombreFuncionario
    End If
End Sub



'======================================================
Private Sub cmdCancel_Click()
'======================================================
  CierraLaBase
  Unload Me
  End
End Sub



'======================================================
Private Sub cmdOK_Click()
'======================================================
    Dim ni As Integer
     'comprobar si la contraseña es correcta
    If txtPassword.Text = "" Then
        txtPassword.SetFocus
    'esta todo bien
    ElseIf UCase(txtPassword.Text) = UCase(vptFuncPass) Then
        'solo para mi
        If vpnFuncionario = 30 Then
            'comando mio (solo mio) permite mostrar el formulario frmMisInformes
            MDIingreso.mInfoAdm.Enabled = True
            MDIingreso.mInfoAdm.Visible = True
        End If
        HistoriaEntra
        Me.Hide
        Unload frmjLogin   'No lo unload para mantener las variables
    'no puede entrar
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        'SendKeys "{HOME}+{END}"
        cuenta = cuenta + 1
        If cuenta > 2 Then
            CierraLaBase
            End
        End If
    End If
End Sub




'======================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
'======================================================
   If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"        ' COMO SI PULSARA ENTER
    End If
End Sub


'======================================================
Public Sub CierraLaBase()
'======================================================
'If adoms.State = adStateOpen Then
'adoms.Close
'Set adoms = Nothing
'End If
adoConn.Close
Set adoConn = Nothing
End Sub
