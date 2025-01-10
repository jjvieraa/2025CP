VERSION 5.00
Begin VB.Form fjParametros 
   BackColor       =   &H0000C000&
   Caption         =   "Parámetros."
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "Salir"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   10
      Left            =   2520
      TabIndex        =   21
      Text            =   " "
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   525
      Index           =   9
      Left            =   2520
      MaxLength       =   120
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "fjParametros.frx":0000
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,#0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   2520
      TabIndex        =   17
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   15
      Text            =   " "
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   13
      Text            =   " "
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,#0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      Text            =   " "
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   9
      Text            =   " "
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Text            =   " "
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Text            =   " "
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "%"
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   23
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Recargo:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3990
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Aviso:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   3390
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Ayuda Social:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Cuota S. Cooperativo:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2670
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Cuota Socio Honorario:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2310
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Cuota Socio Activo:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1950
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Último Nro. Recibo:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Firma:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Último Nro. Orden:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   870
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Dia de Vencimiento:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Presupuesto:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "fjParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoParam As New ADODB.Recordset
Dim adoAcc As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

'========================
Private Sub Form_Load()
'========================
On Error GoTo jb001
    Dim nM As Integer
        
    Text1(0).MaxLength = 6
    Text1(3).MaxLength = 25
    Text1(9).MaxLength = 120
    
    adoAcc.Open "SELECT * FROM TBL_Accion", adoconn, adOpenDynamic, adLockOptimistic, adCmdText
    adoParam.Open "SELECT * FROM TBL_Parametros", adoconn, adOpenDynamic, adLockOptimistic, adCmdText
    adoParam.MoveFirst
    
    'coloca los campos
    For nM = 0 To 10
        Select Case nM
            Case 5, 6, 7, 8, 10
                Text1(nM).Text = Format(adoParam(nM), "#,#0.00")
            Case Else
                Text1(nM).Text = adoParam(nM)
            End Select
    Next
Exit Sub
jb001:
MsgBox ("ERROR jb001: " & Err.Description)
End
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim nM As Integer
    
      
    
    'Guarda los campos
    For nM = 0 To 10
        Select Case nM
            Case 5, 6, 7, 8, 10
                adoParam(nM) = CSng(Text1(nM).Text)
            Case Else
                adoParam(nM) = Text1(nM).Text
            End Select
    Next
    adoParam.Update
  
    'Si hubo cambios guarda en tbl_accion
    Call fHuboCambios
    
    'Cierra la tabla
    adoParam.Close
    Set adoParam = Nothing
    adoAcc.Close
    Set adoAcc = Nothing
    
    'Toma otra ves los parametros
    If Not fTomaMesOperacYParametros() Then
        MsgBox "Error 3423: Al tomar parámetros", vbCritical, "Atención"
        End
    End If
End Sub


Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
        Select Case Index
            Case 5, 6, 7, 8, 10             'single
                If Not IsNumeric(Text1(Index).Text) Then Cancel = True
                Text1(Index).Text = Format(Text1(Index).Text, "#,#0.00")
            Case 0                          'aaaamm presupuesto
                If Not IsNumeric(Text1(Index).Text) Then Cancel = True
                Dim nMes As Byte
                nMes = CByte(Right(Text1(Index).Text, 2))
                If nMes > 12 Or nMes < 0 Then Cancel = True
                Dim nAño As Integer
                nAño = CInt(Left(Text1(Index).Text, 4))
                If nAño > 2050 Or nAño < 2000 Then Cancel = True
        End Select

End Sub

'si hubo cambio de parametros
'lo guarda en tbl_accion
Private Sub fHuboCambios()
    If Not vptMesPresup = adoParam("prm_prspst") Then
        fCmbPrm "vptMesPresup", vptMesPresup, adoParam("prm_prspst")
    End If
    If Not vpnPrspHst = adoParam("PRM_PRSPHST") Then
        fCmbPrm "vpnPrspHst", vpnPrspHst, adoParam("PRM_PRSPHST")
    End If
    If Not vplNroOrden = adoParam("NroOrden") Then
        fCmbPrm "vplNroOrden", vplNroOrden, adoParam("NroOrden")
    End If
    If Not vplNroRecibo = adoParam("prm_NroRecibo") Then
        fCmbPrm "vplNroRecibo", vplNroRecibo, adoParam("prm_NroRecibo")
    End If
    If Not vpsCuotaSAct = adoParam("prm_SocAct") Then
        fCmbPrm "vpsCuotaSAct", vpsCuotaSAct, adoParam("prm_SocAct")
    End If
    If Not vpsCuotaSHon = adoParam("prm_SocHon") Then
        fCmbPrm "vpsCuotaSHon", vpsCuotaSHon, adoParam("prm_SocHon")
    End If
    
    If Not vpsCuotaSCop = adoParam("prm_SocCop") Then
        fCmbPrm "vpsCuotaSCop", vpsCuotaSCop, adoParam("prm_SocCop")
    End If
    If Not vpsAyuda = adoParam("prm_Ayuda") Then
        fCmbPrm "vpsAyuda", vpsAyuda, adoParam("prm_Ayuda")
    End If
     If Not vpsRecargo = adoParam("prm_recarg") Then
        fCmbPrm "vpsRecargo", vpsRecargo, CStr(adoParam("prm_recarg"))
    End If
End Sub

Private Sub fCmbPrm(sPrm As String, sP1 As Variant, sP2 As Variant)
adoAcc.AddNew
adoAcc("acc_accion") = 30
adoAcc("acc_NroIdent") = sPrm
adoAcc("acc_Detalle") = "Cambio en Parametros"
adoAcc("acc_Text1") = sP1
adoAcc("acc_Text2") = sP2
adoAcc("acc_Func") = vpnFuncionario
adoAcc("acc_FDia") = Format(Date, "short date")
adoAcc("acc_FHora") = Format(Time, "short time")
adoAcc.Update
End Sub
