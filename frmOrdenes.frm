VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjOrdenes 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Orden"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImpr 
      BackColor       =   &H0080C0FF&
      Caption         =   "Imprimir:"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H0080C0FF&
      Caption         =   "Guardar"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   1515
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   7095
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   5400
         TabIndex        =   8
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   5400
         TabIndex        =   7
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   5
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   600
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   840
         MaxLength       =   2
         TabIndex        =   4
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   12
         TabIndex        =   3
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   660
         Width           =   1230
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   3
         TabIndex        =   2
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   3360
         TabIndex        =   48
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   5400
         TabIndex        =   47
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "No.Orden:"
         Height          =   210
         Left            =   4440
         TabIndex        =   46
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cuota:"
         Height          =   210
         Left            =   4440
         TabIndex        =   45
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pago:"
         Height          =   210
         Left            =   4440
         TabIndex        =   44
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fecha:"
         Height          =   210
         Left            =   2160
         TabIndex        =   43
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mon:"
         Height          =   210
         Left            =   2160
         TabIndex        =   42
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Plan:"
         Height          =   210
         Left            =   120
         TabIndex        =   41
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Valor:"
         Height          =   210
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Comercio:"
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   5
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   2685
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6585
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1395
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   2461
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   3270
      Left            =   120
      TabIndex        =   11
      Top             =   60
      Width           =   6975
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "en"
         ToolTipText     =   "F2=Dependiente"
         Top             =   1275
         Width           =   1230
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   16
         Left            =   5280
         TabIndex        =   54
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   15
         Left            =   5400
         TabIndex        =   53
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Límite:"
         Height          =   255
         Left            =   4320
         TabIndex        =   52
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Autorizado:"
         Height          =   255
         Left            =   4320
         TabIndex        =   51
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cuota:"
         Height          =   225
         Left            =   120
         TabIndex        =   50
         Top             =   2700
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   1320
         TabIndex        =   49
         Top             =   2680
         Width           =   1005
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   6960
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Docum:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1815
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   11
         Left            =   1320
         TabIndex        =   37
         Top             =   1860
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderStyle     =   6  'Inside Solid
         X1              =   0
         X2              =   6960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   10
         Left            =   1200
         TabIndex        =   36
         Top             =   900
         Width           =   4575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Direc:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   705
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   9
         Left            =   1200
         TabIndex        =   34
         Top             =   660
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Docum:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   705
         Width           =   705
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   6
         Left            =   1320
         TabIndex        =   30
         Top             =   1620
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nombr.Depend."
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1590
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nro.Depend."
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1305
         Width           =   1050
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Disponible:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2925
         Width           =   960
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenes:"
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   2475
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   1320
         TabIndex        =   23
         Top             =   2940
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1320
         TabIndex        =   22
         Top             =   2465
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1320
         TabIndex        =   21
         Top             =   2250
         Width           =   1005
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1200
         TabIndex        =   20
         Top             =   465
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Saldo:"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2265
         Width           =   825
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   3375
         TabIndex        =   15
         Top             =   195
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "No.Cobro:"
         Height          =   345
         Left            =   2520
         TabIndex        =   14
         Top             =   195
         Width           =   840
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   440
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nro.Socio:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0E0FF&
      Caption         =   "lbl"
      Height          =   225
      Index           =   8
      Left            =   3570
      TabIndex        =   32
      Top             =   3360
      Width           =   3450
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0E0FF&
      Caption         =   "lbl"
      Height          =   225
      Index           =   7
      Left            =   60
      TabIndex        =   31
      Top             =   3360
      Width           =   3450
   End
   Begin VB.Label lblMensajes 
      BackColor       =   &H00C0E0FF&
      Caption         =   "12345678901234567890123456789012345678901234567890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3330
      TabIndex        =   17
      Top             =   6630
      Width           =   3585
   End
End
Attribute VB_Name = "fjOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Option Explicit
'#Const kCasa = -1
Dim adoM As New ADODB.Recordset
Dim adoNumOrden As New ADODB.Recordset 'toma el No Orden numero de orden
'Dim adoB As New ADODB.Recordset 'que no este repetido

Dim cSOC As New clsSocios
'Dim cOrd As New clsOrdenes LA UTILIZO GLOBAL
Dim cCom As New clsComercios
Dim cDEP As New clsDepend
Dim cTC As New clsTCambio
Dim sTsCb As Single     'tasa cambio

Dim sMome1(25) As String
Dim sMome2(25) As String
'Dim sMeses1 As String       'para kTipoRecibo = 2
'Dim sMeses2 As String       'para kTipoRecibo = 2
'Dim sMeses As String        'para kTipoRecibo = 2

'LBL
Const kNroCob = 0
Const kNombre = 1
Const kSaldoSueldo = 2
Const kOrdenes = 3
Const kDisponible = 4
Const kComNmb = 5
Const kDepNmb = 6
Const kFechaVtoPresup = 7
Const kOrdnsSigPres = 8
Const kSocDoc = 9
Const kSocDir = 10
Const kDepDoc = 11
Const kNOrden = 12
Const kMon = 13
Const kCuotEnP = 14
Const kDepAutorizado = 15
Const kDepLimite = 16


'TXT
Const kNroSoc = 0
Const kNroCom = 1
Const kNroDep = 2
Const kValor = 3
Const kPlan = 4
Const kMoneda = 5
Const kCuota = 6
Const kPago = 7


Const kTipoImpr = 1     '1=print 2=drOrden


'=====================================================================
Private Sub Form_Load()
'=====================================================================
MDIingreso.StatusBar1.Panels(2).Text = "Inicializa etiquetas..."
Call InicializaEtiquetas
'MDIingreso.StatusBar1.Panels(2).Text = "Abre Parametros..."
'Call AbreParametrosParaTomarNoOrden
MDIingreso.StatusBar1.Panels(2).Text = "Inicia socios..."
cSOC.msInicia
'muestra el boton imprimir
'solo como una prueba
If vpnFuncionario = kYO Then
    cmdImpr.Visible = True
Else
    cmdImpr.Visible = False
End If
cmdGrabar.Enabled = False
'fjAviso3.Hide
MDIingreso.StatusBar1.Panels(2).Text = ""
End Sub





'=====================================================================
Private Sub InicializaEtiquetas()
'=====================================================================
    Dim ni As Byte
    
    For ni = 0 To 16
        lbl(ni).Caption = ""
    Next
    For ni = 2 To 4
        lbl(ni).Caption = "0"
    Next
    lbl(14).Caption = "0"
    lblMensajes.Caption = ""
    Set DataGrid1.DataSource = Nothing
    
    For ni = 0 To 7
        txt(ni).Text = ""
    Next
    txt(kValor).Text = "0,00"
    mskFecha.Mask = ""
    mskFecha.Text = ""
    mskFecha.Mask = "##/##/####"
    lbl(kDisponible).ForeColor = &H80000012         'negro
    'DataGrid1.Visible = False
    cmdGrabar.Enabled = False
    cmdImpr.Enabled = False
    txt(kMoneda).Text = "P"
End Sub












'====================================
Private Sub cmdGrabar_Click()
'====================================
Dim lmome As Long       ' el No de orden

lblMensajes.Caption = "7.Grabando...."
lblMensajes.Refresh

cmdGrabar.Enabled = False
Screen.MousePointer = vbHourglass

' Guarda el No de Orden
lmome = mGlob.cOrd.TomaYGrabaNumOrden               'TomaYGrabaNumOrden desde tbl_parametros
If lmome = 0 Then                                                                ' hubo un error en el No de orden
    MouseOn
    MsgBox "Error 3354a: No tomó el Número de Orden", vbCritical, "Anulando Operación"
    Exit Sub
End If
mGlob.cOrd.vlNroOrden = lmome
lbl(kNOrden).Caption = lmome
lbl(kNOrden).Refresh
'coloca en la clase el Número de la Orden
mGlob.cOrd.vlNroOrden = CLng(lbl(kNOrden).Caption)

lblMensajes.Caption = "6.Cerando....."
lblMensajes.Refresh

'inicializa las variables sMome1 y sMome2
'para LlenaSMeses llamada desde ImprimeUnaOrden
Dim nk As Integer
For nk = 0 To 24
    sMome1(nk) = ""
    sMome2(nk) = ""
Next

'Por las dudas toma otra vez el No de orden
'lbl(kNOrden).Caption = cOrd.mfTomaNumOrden
'cOrd.vlNroOrden = lbl(kNOrden).Caption

'imprime
lblMensajes.Caption = "5.Controlando No...."
lblMensajes.Refresh
Call ImprimeUnaOrden           '<<----------------------

'graba
lblMensajes.Caption = "4.Abriendo..."
lblMensajes.Refresh
'cOrd.msInicia2

lblMensajes.Caption = "3.Guardando..."
lblMensajes.Refresh
cOrd.GrabaEnAdoOrdns

lblMensajes.Caption = "2.Cerrando..."
lblMensajes.Refresh
'cOrd.msTermina

'graba No Orden
lblMensajes.Caption = "1.No Orden..."
lblMensajes.Refresh
'cOrd.msGrabaNumOrden

'final
InicializaEtiquetas
lblMensajes.Caption = ""
lblMensajes.Refresh
DoEvents
Screen.MousePointer = vbDefault
txt(kNroSoc).SetFocus

End Sub









'====================================
Private Sub GuardaValores()
'====================================
cOrd.vlNroSoc = CLng(txt(kNroSoc).Text)
cOrd.vlNroComerc = CLng(txt(kNroCom).Text)
cOrd.vlNroDepend = CLng(txt(kNroDep).Text)
cOrd.vnsCuota = CSng(lbl(kCuotEnP).Caption)
cOrd.vdFEmis = CDate(mskFecha.Text)
cOrd.vdFVto = CDate("10/" & txt(kPago).Text)
cOrd.vnPlan = CInt(txt(kPlan).Text)
cOrd.vnCtasPaga = 0
cOrd.vnsEntCta = 0
cOrd.vnsRecarg = 0
cOrd.vsMoneda = txt(kMoneda).Text
cOrd.vnsMECuota = CSng(txt(kCuota).Text)
cOrd.vsFunc = vpnFuncionario
cOrd.vsFDia = Format(Date, "short date")
cOrd.vsFHora = Format(Time, "short time")
cOrd.vdCerro = CDate(0)
End Sub







'IMPR====================================
Private Sub ImprimeUnaOrden()
'IMPRIME ORDEN
'====================================
DoEvents
If cCom.nbCantRec = 0 Then
    Exit Sub
End If
lblMensajes.Caption = "Imprimiendo..."
LLenaSMeses

'drOrden.Show
Dim ni As Integer
Dim Sign1 As String
Dim Sign2 As String

'papel fanfold 12 pulgadas................
'#If kCasa Then
'msImpresoraDeterminadaGenerica
'#End If
'Printer.PaperSize = vbPRPSFanfoldStdGerman

'crea el ado..............................
' Es unpequeño ado con los encabezado y pie de recibos
Set adoM = New ADODB.Recordset
adoM.Fields.Append "C1", adChar, 600
adoM.Fields.Append "C2", adChar, 20
adoM.Fields.Append "C3", adChar, 50
adoM.Fields.Append "c4", adChar, 50
adoM.Open
Select Case txt(5).Text
    Case "P"
        Sign1 = " $"
        Sign2 = "PESOS"
    Case "D"
        Sign1 = "U$"
        Sign2 = "DOLARES"
    Case "R"
        Sign1 = " R"
        Sign2 = "Real"
    Case "A"
        Sign1 = "$A"
        Sign2 = "P.Argentinos"
End Select
Dim sMomento As String
Dim sMomento1 As String
    
     sMomento1 = "ORDEN Nro: " & lbl(12).Caption & _
             Space(15) & "Rivera, " & mskFecha.Text
     sMomento = sMomento1
     sMomento1 = mfRepite("=", 49)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("En " & Sign2, 50) & sMome1(0) & " " & _
            sMome2(0) & "  " & sMome1(8) & " " & sMome2(8) & "  " & sMome1(16) & " " & sMome2(17)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("Sres: ", 50) & sMome1(1) & " " & _
            sMome2(1) & "  " & sMome1(9) & " " & sMome2(9) & "  " & sMome1(17) & " " & sMome2(18)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("Empresa " & lbl(5).Caption, 50) & sMome1(2) & " " & _
            sMome2(2) & "  " & sMome1(10) & " " & sMome2(10) & "  " & sMome1(18) & " " & sMome2(19)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("De nuestra consideración:", 50) & sMome1(3) & " " & _
            sMome2(3) & "  " & sMome1(11) & " " & sMome2(11) & "  " & sMome1(19) & " " & sMome2(20)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("Por la presente sírvase entregar al Socio " & txt(0).Text, 50) & sMome1(4) & " " & _
            sMome2(4) & "  " & sMome1(12) & " " & sMome2(12) & "  " & sMome1(20) & " " & sMome2(21)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("Sr. " & mfCompleta(lbl(kNombre).Caption, 26) & " CI: " & lbl(9).Caption, 50) & sMome1(5) & " " & _
            sMome2(5) & "  " & sMome1(13) & " " & sMome2(13) & "  " & sMome1(21) & " " & sMome2(22)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("Mercadería por el importe de  " & Sign1 & "   " & txt(3).Text, 50) & sMome1(6) & " " & _
            sMome2(6) & "  " & sMome1(14) & " " & sMome2(14) & "  " & sMome1(22) & " " & sMome2(23)
     sMomento = sMomento & vbCrLf & sMomento1
     sMomento1 = mfCompleta("   ", 50) & sMome1(7) & " " & _
            sMome2(6) & "  " & sMome1(15) & " " & sMome2(14) & "  " & sMome1(23) & " " & sMome2(23)
     sMomento = sMomento & vbCrLf & sMomento1
     'MsgBox Len(sMomento)


For ni = 1 To cCom.nbCantRec
   adoM.AddNew
   
    adoM!c1 = sMomento
    Select Case ni
        Case 1
            adoM!c2 = "1) VIA CLIENTE"
            adoM!c3 = "Sin otro particular saluda atte.:"
            adoM!c4 = "p/Circ.Policial " & vptNombreFuncionario
        Case 2
            adoM!c2 = "2) VIA COPIA"
            adoM!c3 = ""
            adoM!c4 = "* * NO VALIDO PARA EL COMERCIO * * "
       Case 3
            adoM!c2 = "3) VIA ARCHIVO"
            adoM!c3 = ""
            adoM!c4 = "SOCIO:" & lbl(kNombre).Caption
    End Select
    adoM.Update
Next ni

'IMPRIMIENDO .........................
If kTipoImpr = 1 Then

    'Printer.ScaleLeft = -0.2 * 1440    MAGEN IZQ
    Printer.ScaleLeft = -0# * 1440
    Printer.ScaleTop = -0
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    
    adoM.MoveFirst
    For ni = 1 To cCom.nbCantRec
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print adoM(0)
        Printer.Print
        Printer.Print
        'Printer.Print
        Printer.Print
        Printer.Print adoM(2)
        Printer.Print Space(30) & adoM(3)
        Printer.Print
        Printer.Print
        Printer.Print Space(5) & adoM(1)
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
       adoM.MoveNext
    Next ni
    Printer.EndDoc
Else
    Set drOrden.DataSource = adoM
    drOrden.Sections(1).Controls(1).DataField = adoM(0).Name
    drOrden.Sections(1).Controls(2).DataField = adoM(1).Name
    drOrden.Sections(1).Controls(3).DataField = adoM(2).Name
    drOrden.Sections(1).Controls(4).DataField = adoM(3).Name
    'drOrden.Sections(1).Controls(6).Caption = sMeses1
    'drOrden.Sections(1).Controls(7).Caption = sMeses2
    'drOrden.Title = sMeses
    
    drOrden.Refresh
    drOrden.Refresh
    drOrden.Show vbModal
    'drOrden.PrintForm
End If
adoM.Close
Set adoM = Nothing
'#If kCasa Then
' msImpresoraDeterminadaStylus
'#End If
lblMensajes.Caption = ""
End Sub






'====================================
Private Sub LLenaSMeses()
'====================================
' La cantidad de cuotas que va a pagar
Dim nFin As Byte
Dim ni As Byte
Dim n2 As Integer
Dim nT As Byte
Dim sMone As String

Select Case cOrd.vsMoneda
    Case "P"
        sMone = "$"
    Case "D"
        sMone = "U$"
    Case "R"
        sMone = "R"
    Case "A"
        sMone = "$A"
End Select

If cOrd.vnPlan > 24 Then
    nFin = 24
Else
    nFin = cOrd.vnPlan
End If
n2 = Month(cOrd.vdFVto)
For ni = 0 To nFin - 1
    nT = n2 + ni
    If nT > 12 Then
        n2 = -ni + 1
        nT = 1
    End If
    sMome1(ni) = mfDevMes(nT)
    'sMome2(nI) = sMone & " " & cORD.vnsMECuota
    sMome2(ni) = Format(cOrd.vnsMECuota, "#,#0.00")
Next ni
'If kTipoImpr = 2 Then
        'sMeses = sMome(0) & " " & sMome(8) & " " & sMome(16) & _
        '        vbCrLf & sMome(1) & " " & sMome(9) & " " & sMome(17) & _
        '        vbCrLf & sMome(2) & " " & sMome(10) & " " & sMome(18) & _
        '        vbCrLf & sMome(3) & " " & sMome(11) & " " & sMome(19) & _
        '        vbCrLf & sMome(4) & " " & sMome(12) & " " & sMome(20) & _
        '        vbCrLf & sMome(5) & " " & sMome(13) & " " & sMome(21) & _
        '        vbCrLf & sMome(6) & " " & sMome(14) & " " & sMome(22) & _
        '        vbCrLf & sMome(7) & " " & sMome(15) & " " & sMome(23)
        
'        sMeses1 = sMome1(0) & _
'                vbCrLf & sMome1(1) & _
'                vbCrLf & sMome1(2) & _
'                vbCrLf & sMome1(3) & _
'                vbCrLf & sMome1(4) & _
'                vbCrLf & sMome1(5) & _
'                vbCrLf & sMome1(6) & _
'                vbCrLf & sMome1(7)
'        sMeses2 = sMome2(0) & _
'                vbCrLf & sMome2(1) & _
'                vbCrLf & sMome2(2) & _
'                vbCrLf & sMome2(3) & _
'                vbCrLf & sMome2(4) & _
'                vbCrLf & sMome2(5) & _
'                vbCrLf & sMome2(6) & _
'                vbCrLf & sMome2(7)
'End If
End Sub





'====================================
Private Sub Recibo1(sPrm As String, sPrm1 As String)
'====================================
Printer.Print "ORDEN Nro. " & Format(sTr(cOrd.vlNroOrden), "######") & _
    Space(30) & "Rivera, " & cOrd.vdFEmis

Printer.Print mfCompleta("Sres.", 50) & "|"
Printer.Print mfCompleta("Empresa " & lbl(5).Caption, 50) & "|"
Printer.Print mfCompleta("De nuestra consideración:", 50) & "|"
Printer.Print mfCompleta("Por la presente sírvase entregar al Socio Nro.:" & cOrd.vlNroSoc, 50) & "|"
Printer.Print mfCompleta("Señor: " & lbl(kNombre).Caption & " CI:" & lbl(kSocDoc).Caption, 50) & "|"
Printer.Print mfCompleta("Mercadería por ", 50) & "|"
Printer.Print "Sin otro particular saluda atte.:"
Printer.EndDoc
End Sub





Private Sub cmdImpr_Click()
ImprimeUnaOrden
End Sub




'=====================================================================
Private Sub cmdSalir_Click()
'=====================================================================
Unload Me

End Sub






'=====================================================================
Private Sub Calculos(nPrm As Byte)
'=====================================================================
'nPrm=1 calcula CUOTA en funcion del valor
'nPrm=2 calcula VALOR en funcion de la cuota
      
    Dim sValor   As Single
    Dim sCuota As Single
    Dim ndia As Byte    'ultimo dia de presupuesto
    Dim dFech As Date
    

    If nPrm = 1 Then
        'si estan todos los datos
        If txt(kValor).Text = "" Or _
            txt(kPlan).Text = "" Or _
            txt(kMoneda).Text = "" Then
                Exit Sub
        End If
        
        'Cuota
        sValor = CSng(txt(kValor).Text)
        sCuota = sValor / CInt(txt(kPlan).Text)
        txt(kCuota).Text = Format(sCuota, "#,#0.00")
    
    
    Else        'NPRM = 2
            'si estan todos los datos
        If txt(kCuota).Text = "" Or _
            txt(kPlan).Text = "" Or _
            txt(kMoneda).Text = "" Then
                Exit Sub
        End If
         'Cuota
         sCuota = CSng(txt(kCuota).Text)
        sValor = sCuota * CInt(txt(kPlan).Text)
        txt(kValor).Text = Format(sValor, "#,#0.00")
    
    End If
       
     '0000000000000000000000 MONEDA EXTRANJ
    If Not txt(kMoneda).Text = "P" Then
        sTsCb = cTC.mfDevuelveCambio(txt(kMoneda).Text, mskFecha.Text)
        sCuota = sCuota * sTsCb
    Else
        sTsCb = 1
    End If

    'calcula que la cuota se pueda pagar
    '6bis)CONTROLA QUE QUEDE SALDO FAVORABLE
    Dim lmome As Long
    lmome = CLng(lbl(kSaldoSueldo).Caption) - CLng(lbl(kOrdenes).Caption) - sCuota
    If lmome < 0 Then
        lbl(kDisponible).ForeColor = &HFF&
    Else
        lbl(kDisponible).ForeColor = &H80000012
    End If
    lbl(kCuotEnP).Caption = Format(sCuota, "#,#0.00")
    lbl(kDisponible).Caption = Format(lmome, "#,#0.00")

    
End Sub






'=====================================================================
Private Sub Form_Unload(Cancel As Integer)
'=====================================================================
cSOC.msTermina
Set cSOC = Nothing
Set fjOrdenes = Nothing
Set cCom = Nothing
Set cOrd = Nothing
Set cDEP = Nothing
Set cTC = Nothing

    'If adoB.State = adStateOpen Then adoB.Close
    'Set adoB = Nothing
    If adoNumOrden.State = adStateOpen Then adoNumOrden.Close
    Set adoNumOrden = Nothing

End Sub






Private Sub mskFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub





'====================================
Private Sub mskFecha_Validate(Cancel As Boolean)
'====================================
If Not IsDate(mskFecha.Text) Then
    Cancel = True
End If
End Sub





'=====================================================================
Private Sub txt_GotFocus(Index As Integer)
'=====================================================================
    Select Case Index
        Case kNroSoc
            If Not vpbVieneDeMuestraTabla Then
                lblMensajes.Caption = ""
                Call InicializaEtiquetas
            End If
            vpbVieneDeMuestraTabla = False
        Case kNroDep
            txt(kNroDep).Text = "0"
        Case kValor
            lblMensajes.Caption = "Por el valor total de la orden"
        Case kPlan
            If txt(kPlan).Text = "" Then
                txt(kPlan).Text = " 1"
            End If
            lblMensajes.Caption = "Cantidad de cuotas de la orden"
        Case kMoneda
            If txt(kMoneda).Text = "" Then
                txt(kMoneda).Text = "P"
            End If
            lblMensajes.Caption = "Monedas: P=pesos R=Real D=Dolar A=P.Argent U=UR"
        Case kPago
            lblMensajes.Caption = "Formato mm/aaaa"
        Case Else
            lblMensajes.Caption = ""
    End Select
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index).Text)
End Sub


'=====================================================================
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'=====================================================================
    Select Case Index
        Case kNroSoc
            If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrSocAlf3
                fjMuestraTabla.Show
            ElseIf KeyCode = 114 Then   'F3
                vpMuestraTabla = kMstrSocPorCobr1    'Por No Cobro
                fjMuestraTabla.Show
            End If
        Case kNroDep
            If KeyCode = vbKeyF2 Then    'f2
                vpMuestraTabla = kMuestraDepend 'pasa la tabla que va a mostrar
                vplNroSocio = CLng(txt(kNroSoc).Text)     'pasa el numero de socio
                fjMuestraTabla.Show
            ElseIf KeyCode = vbKeyEscape Then
                txt(kNroSoc).SetFocus
            End If
        Case kNroCom
            If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMuestraComercios
                fjMuestraTabla.Show
            ElseIf KeyCode = vbKeyEscape Then
                txt(kNroDep).SetFocus
            End If
    End Select
End Sub





'=====================================================================
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'=====================================================================
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub





'=====================================================================
Private Sub txt_LostFocus(Index As Integer)
'=====================================================================

lblMensajes.Caption = ""
Select Case Index
    Case 3          'valor
        Calculos (1)
    Case 4          'plan
        Calculos (1)
    Case 5          'moneda
        Calculos (1)
        BuscaDatosParte4
    Case 6          'cuota
        Calculos (2)
        BuscaDatosParte4
End Select
End Sub





'=====================================================================
Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'=====================================================================
    
    Select Case Index
        Case kNroSoc
            'tiene que ser numerico distinto de cero
            If Not IsNumeric(txt(kNroSoc).Text) Then
                Cancel = True
            ElseIf CInt(txt(kNroSoc).Text) = 0 Then
                Cancel = True
            ' Toma datos del cliente
            ElseIf Not BuscaDatosParte1 Then
                    Cancel = True
            End If
        Case kNroDep
            'DEPENDIENTE PUEDE SER CERO
            If Not IsNumeric(txt(kNroDep).Text) Then
                Cancel = True
            ElseIf Not BuscaDatosParte2 Then
                Cancel = True
            End If
        Case kNroCom
            If Not IsNumeric(txt(kNroCom).Text) Then
                Cancel = True
            ElseIf CInt(txt(kNroCom).Text) = 0 Then
                Cancel = True
            ElseIf Not BuscaDatosParte3 Then
                Cancel = True
            End If
        Case kValor
            If Not IsNumeric(txt(kValor).Text) Then
                Cancel = True
            Else
                txt(kValor).Text = mfCambiaPuntoPorComa(txt(kValor).Text)
                txt(kValor).Text = Format(txt(kValor).Text, "#,#0.00")
            End If
        Case kPlan
            If Not IsNumeric(txt(kPlan).Text) Then
                Cancel = True
            End If
        Case kMoneda
            txt(kMoneda).Text = UCase(txt(kMoneda).Text)
            Select Case txt(kMoneda).Text
                Case "P"
                    lbl(kMon).Caption = "Pesos"
                Case "D"
                    lbl(kMon).Caption = "Dolar"
               Case "R"
                    lbl(kMon).Caption = "Real "
                Case "A"
                    lbl(kMon).Caption = "P.Arg"
                Case "U"
                    lbl(kMon).Caption = "U.R. "
                Case Else
                    lbl(kMon).Caption = ""
                    Cancel = True
            End Select
        Case kPago
            Dim sMome As String
            Dim bMome As Byte
            Dim bMes As Byte
            Dim nAno As Integer
            
            'encuentra dosde esta el /
            sMome = txt(kPago).Text
            If Mid(sMome, 2, 1) = "/" Then
                bMome = 2
            ElseIf Mid(sMome, 3, 1) = "/" Then
                bMome = 3
            Else
                Cancel = True
                Exit Sub
            End If
            bMes = CByte(Left(sMome, bMome - 1))
            nAno = CInt(Right(sMome, 4))
            If Not IsNumeric(bMes) Or bMes > 12 Or bMes < 1 Then
                Cancel = True
            ElseIf Not IsNumeric(nAno) Or nAno > 2020 Or nAno < 1990 Then
                Cancel = True
            End If
        Case kCuota
            'txt(kCuota).Text = mfCambiaPuntoPorComa(txt(kCuota).Text)
      
    End Select
End Sub






'=====================================================================
Private Function BuscaDatosParte1() As Boolean
'=====================================================================
    'BUSCA DATOS DEL SOCIO

    Dim lmome As Long
    
    ' Abre tabla socios y ubica el registro que corresponde
    m3Aviso "Buscando Datos..."
    cSOC.mfAbreTablaSociosOrdenSocio
    cSOC.vlNroSoc = CLng(txt(kNroSoc).Text)
    
    '1) BUSCA AL SOCIO
    If Not cSOC.mfBuscaSocio Then
        MsgBox "4552: Socio no encontrado"
        fColocaMarcaAccion 40, "MalNroSoc", "Mal Numero de Socio " & cSOC.vlNroSoc, "", ""
        BuscaDatosParte1 = False
        Exit Function
    End If
    
    '2) MUESTRA DATOS DEL SOCIO
    lbl(kNombre).Caption = Left(cSOC.vsApellido & " " & cSOC.vsNombre, 40)
    lbl(kNroCob).Caption = cSOC.vsNroCob
    lbl(kSaldoSueldo).Caption = Format(cSOC.vnsLimite, "#,#0.00")
    lbl(kSocDoc).Caption = cSOC.vsCi
    lbl(kSocDir).Caption = cSOC.vsDireccion
    
    '3)BUSCA LAS ORDENES QUE TIENE EL SOCIO
    'cOrd.msInicia
    
    mGlob.cOrd.vlNroSoc = txt(0)
    m3Aviso "Espere: Buscando Ordenes..."
    ' Busca las ordenes en mglob.cord.adoOrdenes
    If Not mGlob.cOrd.fBuscaOrdenesUnSocio Then     'HUBO PROBLEMAS
        'MsgBox "4553: Problemas al Buscar Ordenes"
        'BuscaDatosParte1 = False
        'Exit Function
    End If
    
    ' Cuenta cuantas son las ordenes
    If mGlob.cOrd.adoOrdenes.RecordCount = 0 Then
        'MsgBox "4554: No tiene Ordenes"
        lbl(kOrdenes).Caption = "0"
        lbl(kDisponible).Caption = lbl(kSaldoSueldo).Caption

    Else
        '4.Prepara el ado: mglob.cord.adoM2
        mGlob.cOrd.msPreparaOrdenesAPagarEnAdoM2 (0)
        
        '5) CALCULA LO QUE DEBE PAGAR EL PROXIMO MES
        'm3Aviso "Espere: Calculando Ordenes..."
        Dim sM1 As Single       'total cuotas que vencen en este presup
        Dim sM2 As Single       'total cuotas que vencen en otro presu
        Dim sMome As Single
        sM1 = 0
        sM2 = 0
        If mGlob.cOrd.mfSumaAdoM2(sM1, sM2) Then
            lbl(kOrdenes).Caption = Format(sM1, "#,#0.00")
            lbl(kOrdnsSigPres).Caption = "Proxs.Prsts:" & Format(sM2, "#,#0,00")
            sMome = CLng(lbl(kSaldoSueldo).Caption) - CLng(lbl(kOrdenes).Caption)
            If sMome < 0 Then
                lbl(kDisponible).ForeColor = &HFF&
            Else
                lbl(kDisponible).ForeColor = &H80000012
            End If
            lbl(kDisponible).Caption = Format(sMome, "#,#0.00")


        
        End If
        
        '6) Muestra las ordenes en el DataGrid
        DoEvents
        mGlob.cOrd.msMuestraAdoM2 DataGrid1
        
       
        m3Aviso ("")
    End If
    txt(kNroDep).SetFocus
    
  
    
    BuscaDatosParte1 = True
    lblMensajes.Caption = ""
    cSOC.msTermina
End Function






'=====================================================================
Private Sub m3Aviso(sPrm As String)
'=====================================================================
    lblMensajes.Caption = sPrm
    lblMensajes.Refresh
End Sub







'=====================================================================
Private Function BuscaDatosParte2() As Boolean       ' DATOS DEL DEPENDIENTE
'=====================================================================
'BUSCA DATOS DEL DEPENDIENTE
    If CLng(txt(kNroDep).Text) = 0 Then
        BuscaDatosParte2 = True
        Exit Function
    End If
    cDEP.vlDepNum = CLng(txt(kNroDep).Text)
    cDEP.vlNroSoc = CLng(txt(kNroSoc).Text)
    If Not cDEP.mfAbreTablaDepend Then
        MsgBox "4554: No abre tabla.."
        BuscaDatosParte2 = False
        Exit Function
    End If
    cDEP.fOrdenaAdoPorDepend
    If Not cDEP.mfBuscaDepend Then
        MsgBox "4555: Dependiente no encontrado"
        BuscaDatosParte2 = False
        Exit Function
    End If
    lbl(kDepNmb).Caption = cDEP.vsDepNmb
    lbl(kDepDoc).Caption = cDEP.vsDepDoc
    If cDEP.vbDepAut = True Then
        lbl(kDepAutorizado).Caption = "SI"
    Else
        lbl(kDepAutorizado).Caption = "NO"
    End If
    lbl(kDepLimite).Caption = Format(cDEP.vdDepLim, "#,#0")
  
    Set cDEP = Nothing
     BuscaDatosParte2 = True
End Function






'====================================================
Private Function BuscaDatosParte3() As Boolean
'====================================================
    
    lbl(5).Caption = cCom.BuscaComercio(CLng(txt(kNroCom).Text))
    'Muestra la fecha en que termina el presupuesto actual
    cOrd.fTomaPresupActual
    lbl(kFechaVtoPresup).Caption = "Presup. Hasta: " & cOrd.vdPrsptoTermina
    Dim lmome As Long
    
    '6)CONTROLA QUE QUEDE SALDO FAVORABLE
    lmome = CLng(lbl(kSaldoSueldo).Caption) - CLng(lbl(kOrdenes).Caption)
    If lmome < 0 Then
        lbl(kDisponible).ForeColor = &HFF&
    End If

    lbl(kDisponible).Caption = Format(lmome, "#,#0.00")
    ' Pone la fecha de hoy
    mskFecha.Mask = ""
    mskFecha.Text = ""
    mskFecha.Text = Format(Date, "short date")
    BuscaDatosParte3 = True
End Function







'=====================================================================
Private Function BuscaDatosParte4() As Boolean
'=====================================================================
    Dim sMome As Long
    Dim lmome As Single
    Dim ndia As Byte    'ultimo dia de presupuesto
    Dim dFech As Date
    Dim tMom As String
    
    
   
    
    'presupuesto que se pagar{a
    ndia = cOrd.mfTomaUltDiaPresup
    dFech = CDate(mskFecha.Text)
    'fecha en que vence la 1a cuota
    'si el dia es <= 10 vence el 10/de este mes: ej 5/1 vence 10/1
    'si el dia es > 10 vence el d10/mes siguiente : ej 28/1 vence 10/2
    If Day(dFech) > ndia Then
        If Month(dFech) = 12 Then
            tMom = ndia & "/1/" & Year(dFech) + 1
        Else
            tMom = ndia & "/" & Month(dFech) + 1 & "/" & Year(dFech)
        End If
                
        'dFech = dFech + 30
    Else
        tMom = ndia & "/" & Month(dFech) & "/" & Year(dFech)
    End If
    dFech = CDate(tMom)
   
    txt(kPago).Text = CStr(Month(dFech)) & _
        "/" & CStr(Year(dFech))
    
    GuardaValores
    
    cmdGrabar.Enabled = True
    cmdImpr.Enabled = True
End Function





'=====================================================================
Private Sub PoneTitulos()
'=====================================================================
DataGrid1.Columns(1).Caption = "Cmrc"
DataGrid1.Columns(1).Width = 500
DataGrid1.Columns(2).Caption = "Ordn"
DataGrid1.Columns(2).Width = 500
DataGrid1.Columns(3).Caption = "Depndt"
DataGrid1.Columns(3).Width = 750
DataGrid1.Columns(4).Caption = "Cuota"
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(4).Alignment = dbgRight
DataGrid1.Columns(4).NumberFormat = "#,#0.00"
DataGrid1.Columns(5).Caption = "Emis"
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(6).Caption = "Vto"
DataGrid1.Columns(6).Width = 1000
DataGrid1.Columns(7).Caption = "Pln"
DataGrid1.Columns(7).Width = 600
DataGrid1.Columns(8).Caption = "Pgs"
DataGrid1.Columns(8).Width = 600
DataGrid1.Columns(9).Caption = "Ent Cta"
DataGrid1.Columns(9).Width = 1000
DataGrid1.Columns(9).Alignment = dbgRight
DataGrid1.Columns(9).NumberFormat = "#,#0.00"
DataGrid1.Columns(10).Caption = "Recargos"
DataGrid1.Columns(10).Width = 1000
DataGrid1.Columns(10).Alignment = dbgRight
DataGrid1.Columns(10).NumberFormat = "#,#0.00"
DataGrid1.Columns(11).Caption = "Mnd"
DataGrid1.Columns(11).Width = 500
DataGrid1.Columns(12).Caption = "MECuota"
DataGrid1.Columns(12).Width = 1000
DataGrid1.Columns(12).Alignment = dbgRight
DataGrid1.Columns(12).NumberFormat = "#,#0.00"
DataGrid1.Columns(13).Caption = "MEPagos"
DataGrid1.Columns(13).Width = 1000
DataGrid1.Columns(13).Alignment = dbgRight
DataGrid1.Columns(13).NumberFormat = "#,#0.00"
End Sub





Private Sub msImpresoraDeterminadaGenerica()

Dim X As Printer
For Each X In Printers
    If X.DeviceName Like "Generica" Then
        Set Printer = X
    End If
Next
End Sub





Private Sub msImpresoraDeterminadaStylus()
Dim X As Printer
For Each X In Printers
    If X.DeviceName Like "Epson Stylus 400" Then
        Set Printer = X
    End If
    'Debug.Print X.DeviceName
Next
End Sub






