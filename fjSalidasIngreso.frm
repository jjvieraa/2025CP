VERSION 5.00
Begin VB.Form fjSalidasYEntradas 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ingreso de Salida"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   915
      Index           =   2
      Left            =   1560
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H0080C0FF&
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H0080C0FF&
      Caption         =   "Agregar"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Salidas"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblRubro 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Detalle:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Entradas"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rubro"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "fjSalidasYEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoC As New ADODB.Recordset
Dim adoR As New ADODB.Recordset
Dim kNuevoRubro As Boolean


Private Sub Form_Load()
    Call PreparaCampos
    If vpMuestraTabla = 1 Then      'funcionario
        Label1.Caption = "Fecha:    " & Date
        adoC.Open "select * FROM tbl_Gastos;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        adoR.Open "select * FROM tbl_GastosRubros ORDER BY sRubro;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    Else                            'administrador
        Label1.Caption = "Adm.  Fecha:    " & Date
        adoC.Open "select * FROM tbl_GastosAdm;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        adoR.Open "select * FROM tbl_GastosRubros ORDER BY sRubro;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    End If
End Sub



Private Sub PreparaCampos()
    Text1(0).Text = ""
    Text1(1).Text = ""          'entradas
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""          'salidas
    cmdAgregar.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    Text1(2).Enabled = False
    Text1(3).Visible = False        'esconde rubro
    Text1(4).Enabled = False
End Sub



Private Sub cmdAgregar_Click()
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(4).Text = ""
    cmdAgregar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    lblRubro.Caption = ""
    cmdSalir.Enabled = True
    Text1(0).Enabled = True
    Text1(1).Enabled = True
    Text1(2).Enabled = True
    Text1(4).Enabled = True
    Text1(0).SetFocus
End Sub



Private Sub cmdCancelar_Click()
    Call PreparaCampos
End Sub




Private Sub cmdsalir_Click()
    Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
    adoC.Close
    Set adoC = Nothing
    adoR.Close
    Set adoR = Nothing
End Sub






Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0:
            If Not IsNumeric(Text1(0).Text) Then
                Cancel = True
            Else
                Cancel = False
                BuscaYMuestraRubro
            End If
        Case 1:         'ENTRADAS
            If IsNumeric(Text1(1).Text) Or Len(Text1(1).Text) = 0 Then
                            Text1(1).Text = Format(Text1(1).Text, "#,#0.00")
            Else
                Cancel = True
            End If
         Case 4:        'SALIDAS
           If IsNumeric(Text1(4).Text) Or Len(Text1(4).Text) = 0 Then
                            Text1(4).Text = Format(Text1(4).Text, "#,#0.00")
            Else
                Cancel = True
            End If

        Case 2:
        Case 3:
            If Len(Text1(3).Text) < 1 Then
                Cancel = True
            End If
    End Select

End Sub



Private Sub cmdGrabar_Click()
    'Revisa validaciones
    If Not IsNumeric(Text1(0).Text) Then
        Text1(0).SetFocus
        Exit Sub
    End If
    If kNuevoRubro Then
        If Len(Text1(3).Text) < 1 Then
            Text1(3).SetFocus
            Exit Sub
        Else
            adoR.AddNew
            adoR!sRubro = CLng(Text1(0).Text)
            adoR!sDetRubro = Text1(3).Text
            adoR.Update
            lblRubro.Visible = True
            Text1(3).Visible = False
        End If
    End If
    'HAY ENTRADA Y SALIDA AL MISMO TIEMPO
    If Not CDbl(IIf(Len(Text1(1).Text) = 0, "0", Text1(1).Text)) = 0 And _
       Not CDbl(IIf(Len(Text1(4).Text) = 0, "0", Text1(4).Text)) = 0 Then
        Text1(1).SetFocus
        Exit Sub
    End If
    
    'Graba
    adoC.AddNew
    adoC!sfecha = Date
    adoC!sHora = Time
    adoC!sRubro = CLng(Text1(0).Text)
    'si esta vacio no lo guarda
    If Not Len(Text1(1).Text) = 0 Then
        adoC!sEntrada = CDbl(Format(Text1(1).Text, "0.0"))
    End If
    If Not Len(Text1(4).Text) = 0 Then
        adoC!sSalida = CDbl(Format(Text1(4).Text, "0.0"))
    End If
    adoC!sDetalle = Text1(2).Text
    adoC!sFunc = vpnFuncionario
    adoC.Update
    Label6.Caption = "Registro: " & adoC!sAutonum
    cmdAgregar.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    Text1(2).Enabled = False
    Text1(4).Enabled = False

End Sub


Private Sub BuscaYMuestraRubro()
    adoR.MoveFirst
    adoR.Find "sRubro =" & CLng(Text1(0).Text)
    'No lo encuentra
    If adoR.EOF Then
        Text1(3).Visible = True
        Text1(3).Enabled = True
        lblRubro.Visible = False
        Text1(3).Refresh
        kNuevoRubro = True
    'Lo encuentra
    Else
       lblRubro.Caption = adoR!sDetRubro
       kNuevoRubro = False
    End If
End Sub





Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
            If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrGastos
                fjMuestraTabla.Show
            End If
    End If
End Sub

