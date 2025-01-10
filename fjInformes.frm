VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fjInformes 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Informes"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "fjInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cobros por Comercio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Todos, clasificados"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Recargos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Todos los cobros sin recargos"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cobros(3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Todos, clasificados"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Gastos Adm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Gastos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H000080FF&
      Caption         =   "Admin 1 Dólares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Las emitidas y las NO canceladas, por mes. En Dólares"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H000080FF&
      Caption         =   "Admin 1 Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Las emitidas y las NO canceladas, por mes. En Pesos"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Todos los Socios(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   $"fjInformes.frx":0442
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Resumen Mens. (4) "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Lo que debe cada socio por grupo"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Resumen Mens. (3) "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Solicitudes a Centro discriminadas por grupos."
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Resumen Mens. (2) "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Solicitudes a Jefatura discriminadas por grupos."
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cobros(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Los cobros agrupados por socio"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Historia(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Vence, No Socio, Comercio, Valor, Impago"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Resumen Mensual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Resumen Mensual"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Una Orden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cuenta Socios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cuenta la cantidad de socios por categorías."
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ord. Com (Total Socio)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   27
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Historia Socio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Fecha, Tipo, NSocio, Debe, Haber, Comercio"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Todos los cobros sin recargos"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ordenes Anuladas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Orden"
      Height          =   1215
      Left            =   1680
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cobro"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numérico"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Alfabético"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tipo"
      Height          =   2175
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Todos"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cooperadores"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1488
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Retirados"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1176
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pensionista"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   864
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Comisiones"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   552
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Activo JPR"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdVer 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ver"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Todos los Socios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   $"fjInformes.frx":04F9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Un Socio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Lo que debe, campos: Vence, NoOrden, Comercio, ValorP, ValorD"
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ordenes Comercio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ordenes Emitidas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1395
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "fjInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Informes
'====================
'0. Ordenes emitidas:  TODAS las ordenes que se emitieron en un período,
'====================  inclusive las anuladas (tipo 4)
'                      Muestra:Fecha, No Orde, No Socio, Nombre, Valor
'                      Muestra: Total
'===================
'10.Resumen Mensual:    De tal fecha a tal otra
'===================    Totales por Carnic, Vales, Cuots, Ayuda, Ordenes
'                       Sub totales por Act,Com,ret,Pens,otros
'                       La suma de carn, vales, ordens = ordenes emitidas (listado 0) - ordenes anuladas
'                       Inclusive las anuladas


Option Explicit

Dim adoClie As New ADODB.Recordset 'cliente
Dim adoCom As New ADODB.Recordset 'comercio
Dim adoC As New ADODB.Recordset 'con las ordenes
Dim adoD As New ADODB.Recordset 'con los datos
Dim adoE As New ADODB.Recordset 'con los datos del comercio
Dim adoCmd As New ADODB.Command
Dim cOrd As New clsOrdenes
Dim cPag As New clsPagos
Dim cCom As New clsComercios
Dim cTC As New clsTCambio
Dim cSocio As New clsSocios

Dim sMome1 As Single            'auxiliar para la rutina suma
Dim sMome2 As Single            'auxiliar para la rutina suma

Const kCantBotones = 4

Const kInfoDia = 0
Const kInfoComercio = 1
Const kInfoUnSocio = 2
Const kInfoTodosSOcios = 3
Const kInfoOrdAnuladas = 4
Const kInfoCobros = 5
Const kInfoHistoriaUnSocio = 6
Const kInfoComercio2 = 7     'para un comercio con subtotales por socio
Const kInfoCuentaSocios = 8     'Cuenta cantidad socios por categoria en una fecha
Const kInfoUnaOrden = 9
Const kInfoResumen = 10
Const kInfoHistoriaUnSocio2 = 11    ' Historia solo de ordenes
Const kInfoCobros2 = 12
Const kInfoResumen2 = 13            'solicitado a Jefatura
Const kInfoResumen3 = 14            'solicitado a Centro
Const kInfoResumen4 = 15            'Lo que deben todos los socios por grupo
Const kInfoTodosSocios2 = 16
Const kinfoAdmin1 = 17
Const kInfoAdmin2 = 18
Const kInfoGastos = 19
Const kInfoGastosAdm = 20
Const kInfoCobros3 = 21
Const kInfoRecargos = 22
Const kInfoCobrosPorComercio = 23

Const kTodos = 20           'muestra todos los objetos
Const k3 = 21
Const kNinguno = 22
Const kSoloVerYSalir = 23
Private mInfo As Byte


'variables publicas para informe mensual
Dim sTot1 As Single           'sub totales
Dim sTot2 As Single           'sub totales
Dim sTot3 As Long               'total registros
Dim tT1 As String           'columna No 1 del reporte
Dim tT2 As String          'columna valor pesos
Dim tT3 As String           'columna valor me
Dim tT4 As String           'columna cant registros
Dim nA1 As Long             'cantidad registros
Dim sT1(5, 5, 4) As Single
Dim sT2(5, 5) As Long


'=======================================================
Private Sub cmdSalir_Click()
'=======================================================
Unload Me
End Sub

'=======================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=======================================================
'Dim nI As Integer
'For nI = 0 To kCantBotones - 1
'cmd(nI).BackColor = &HC0E0FF
'Next
End Sub

'=======================================================
Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'=======================================================
'cmd(Index).BackColor = &H80FF&

End Sub

'=======================================================
Private Sub Form_Load()
'=======================================================
    'NIVELES DE SEGURIDAD ................................
    'If vpnFuncionario = 30 Then
    '    cmd(kPrueba).Visible = True
    '    cmd(kPrueba).Enabled = True
    'End If
    PBar.Visible = False
End Sub

'======================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
'======================================================
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"        ' COMO SI PULSARA ENTER
    End If
End Sub


'=======================================================
Private Sub cmd_Click(Index As Integer)
'=======================================================
Select Case Index
'minfo es el parámetro para el comando cmdVER_click
Case kInfoDia       'todas las ordenes
    fjInformes.Caption = "Informe: todas las Órdenes"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kTodos)
    mInfo = kInfoDia

Case kInfoComercio   'ordenes de comercio
    fjInformes.Caption = "Informe: Órdenes de Comercio"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Label4.Caption = "Comercio:"
    Text1.Text = Date
    Text2.Text = Date
    Text3.Text = ""
    mInfo = kInfoComercio
    Text1.TabIndex = 0
    Text2.TabIndex = 1
    Text3.TabIndex = 2
    cmdVer.TabIndex = 3
    cmdSalir.TabIndex = 4
    Text3.ToolTipText = "F2=Alfab"
    Call MuestraObjetos(k3)

Case kInfoComercio2    'ordenes de comercio
    fjInformes.Caption = "Informe: Órdenes de Comercio"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Label4.Caption = "Comercio:"
    Text1.Text = Date
    Text2.Text = Date
    Text3.Text = ""
    mInfo = kInfoComercio2
    Text1.TabIndex = 0
    Text2.TabIndex = 1
    Text3.TabIndex = 2
    cmdVer.TabIndex = 3
    cmdSalir.TabIndex = 4
    Text3.ToolTipText = "F2=Alfab"
    Call MuestraObjetos(k3)

Case kInfoHistoriaUnSocio 'historia de un socio
    fjInformes.Caption = "Informe: Historia de un socio"
    Label1.Caption = "No Socio:"
    Text1.Text = ""
    Text1.TabIndex = 0
    cmdVer.TabIndex = 1
    cmdSalir.TabIndex = 2
    Call MuestraObjetos(kInfoUnSocio)
    Text1.SetFocus
    mInfo = kInfoHistoriaUnSocio
    Text1.ToolTipText = "F2=Alfab F3=N Cob"
    
Case kInfoHistoriaUnSocio2 'historia de un socio
    fjInformes.Caption = "Informe: Historia de un socio, solo órdenes"
    Label1.Caption = "No Socio:"
    Text1.Text = ""
    Text1.TabIndex = 0
    cmdVer.TabIndex = 1
    cmdSalir.TabIndex = 2
    Call MuestraObjetos(kInfoUnSocio)
    Text1.SetFocus
    mInfo = kInfoHistoriaUnSocio2
    Text1.ToolTipText = "F2=Alfab F3=N Cob"
    
Case kInfoUnSocio  'deuda de un socio
    fjInformes.Caption = "Informe: Estado de un socio"
    Label1.Caption = "No Socio:"
    Text1.Text = ""
    Text1.TabIndex = 0
    cmdVer.TabIndex = 1
    cmdSalir.TabIndex = 2
    Call MuestraObjetos(kInfoUnSocio)
    Text1.SetFocus
    mInfo = kInfoUnSocio
    Text1.ToolTipText = "F2=Alfab F3=N Cob"


Case kinfoAdmin1
    fjInformes.Caption = "Informe: Ordenes emitidas y ordenes No Canceladas por mes, en Pesos"
    cmdVer.TabIndex = 0
    cmdSalir.TabIndex = 1
    Call MuestraObjetos(kNinguno)
    Call mInfoAdmin1(1)

Case kInfoAdmin2
    fjInformes.Caption = "Informe: Ordenes emitidas y ordenes No Canceladas por mes, en Dólares"
    cmdVer.TabIndex = 0
    cmdSalir.TabIndex = 1
    Call MuestraObjetos(kNinguno)
    Call mInfoAdmin1(2)


Case kInfoTodosSOcios   'deuda de todos los socios
    fjInformes.Caption = "Informe: Estado de todos los socios"
    Label1.Caption = "Hasta presup (mmaaaa)(0=Actual): "
    Text1.TabIndex = 0
    Frame1.TabIndex = 1
    Frame1.TabIndex = 2
    cmdVer.TabIndex = 3
    cmdSalir.TabIndex = 4
    Call MuestraObjetos(kInfoTodosSOcios)
    mInfo = kInfoTodosSOcios
    Option1(5).Value = True
    Option2(1).Value = True
    Text1.SetFocus
    
Case kInfoTodosSocios2   'deuda de todos los socios
    fjInformes.Caption = "Informe: Estado de todos los socios"
    Label1.Caption = "Hasta presup (mmaaaa)(0=Actual): "
    Text1.TabIndex = 0
    Frame1.TabIndex = 1
    Frame1.TabIndex = 2
    cmdVer.TabIndex = 3
    cmdSalir.TabIndex = 4
    Call MuestraObjetos(kInfoTodosSOcios)
    mInfo = kInfoTodosSocios2
    Option1(5).Value = True
    Option2(1).Value = True
    Text1.SetFocus

Case kInfoOrdAnuladas       'Ordenes anuladas
    fjInformes.Caption = "Informe: Ordenes Anuladas"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kTodos)
    mInfo = kInfoOrdAnuladas
    
Case kInfoResumen
    fjInformes.Caption = "Informe: Resumen Mensual"
    Label1.Caption = "Presupuesto presup (mmaaaa)(0=Actual): "
    Text1.Text = ""
    Text1.TabIndex = 0
    cmdVer.TabIndex = 1
    cmdSalir.TabIndex = 2
    Call MuestraObjetos(kInfoUnSocio)
    Text1.SetFocus
    mInfo = kInfoResumen

Case kInfoResumen2
    fjInformes.Caption = "Informe: Resumen Mensual"
    Label1.Caption = "Presupuesto presup (mmaaaa)(0=Actual): "
    Text1.Text = ""
    Text1.TabIndex = 0
    cmdVer.TabIndex = 1
    cmdSalir.TabIndex = 2
    Call MuestraObjetos(kInfoUnSocio)
    Text1.SetFocus
    mInfo = kInfoResumen2
 
 Case kInfoResumen3
    fjInformes.Caption = "Informe: Resumen Mensual"
    Label1.Caption = "Presupuesto presup (mmaaaa)(0=Actual): "
    Text1.Text = ""
    Text1.TabIndex = 0
    cmdVer.TabIndex = 1
    cmdSalir.TabIndex = 2
    Call MuestraObjetos(kInfoUnSocio)
    Text1.SetFocus
    mInfo = kInfoResumen3
 
 Case kInfoResumen4
    fjInformes.Caption = "Informe: Resumen Mensual"
    cmdVer.TabIndex = 0
    cmdSalir.TabIndex = 1
    Call MuestraObjetos(kNinguno)
    Call mInfoResumen2(3)
    
Case kInfoCobros        'cobros
    fjInformes.Caption = "Informe: Cobros"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kInfoCobros)
    mInfo = kInfoCobros
    
Case kInfoCobros2        'cobros
    fjInformes.Caption = "Informe: Cobros"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kInfoCobros)
    mInfo = kInfoCobros2
    
Case kInfoCobros3        'cobros
    fjInformes.Caption = "Informe: Cobros"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kInfoCobros)
    mInfo = kInfoCobros3
    
Case kInfoCobrosPorComercio    'cobros por comercio
    fjInformes.Caption = "Informe: Cobros por Comercio"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Label4.Caption = "Comercio:"
    Text1.Text = Date
    Text2.Text = Date
    Text3.Text = ""
    mInfo = kInfoCobrosPorComercio
    Text1.TabIndex = 0
    Text2.TabIndex = 1
    Text3.TabIndex = 2
    cmdVer.TabIndex = 3
    cmdSalir.TabIndex = 4
    Text3.ToolTipText = "F2=Alfab"
    Call MuestraObjetos(k3)
    
Case kInfoRecargos        'Recargos
    fjInformes.Caption = "Informe: Recargos Financieros"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kInfoCobros)
    mInfo = kInfoRecargos

Case kInfoCuentaSocios
    fjInformes.Caption = "Informe: Cuenta Socios por Tipo"
    Label1.Caption = "Ingresados hasta el:"
    Text1.Text = Date
    Call MuestraObjetos(kInfoCuentaSocios)
    mInfo = kInfoCuentaSocios
    
Case kInfoUnaOrden
    Call MuestraObjetos(kNinguno)
    fjInformes.Caption = "Informe: Mira una Orden"
    fjPideDato.Caption = "Nro de Orden:"
    fjPideDato.Show vbModal
    If vpbCancel = False Then
        vpFormMovim = kFormMira
        fjOrden2.Show
    End If

Case kInfoGastos       'Gastos
    fjInformes.Caption = "Informe: Gastos"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kInfoGastos)
    mInfo = kInfoGastos

Case kInfoGastosAdm       'Gastos
    If Not vpnFuncionario = 5 Then Exit Sub
    fjInformes.Caption = "Informe: Gastos"
    Label1.Caption = "Informe desde el:"
    Label2.Caption = "Hasta el:"
    Text1.Text = Date
    Text2.Text = Date
    Call MuestraObjetos(kInfoGastosAdm)
    mInfo = kInfoGastosAdm
    
End Select
CierraTodo
End Sub

'=======================================================
Private Sub MuestraObjetos(nPrm As Byte)
'=======================================================
Select Case nPrm
    Case kNinguno
        Muestra2Objetos False, False, False, False, False, False, _
         False, False, False, False, False

    Case kTodos
          Muestra2Objetos True, True, True, True, True, False, _
         True, True, False, False, False
         
   
    Case kInfoCobros, kInfoGastos, kInfoGastosAdm
          Muestra2Objetos True, True, True, True, True, False, _
         True, False, False, False, False
   
    Case k3
         Muestra2Objetos True, True, True, True, True, True, _
         True, True, False, False, True
    
    Case kInfoUnSocio
         Muestra2Objetos True, False, True, True, False, False, _
         True, False, False, False, False
 
    Case kInfoTodosSOcios
         Muestra2Objetos True, False, True, True, False, False, _
         True, False, True, True, False
    
    Case kInfoCuentaSocios
         Muestra2Objetos True, False, False, True, False, False, _
         True, False, False, False, False
    Case kSoloVerYSalir
         Muestra2Objetos False, False, False, False, False, False, _
         True, False, False, False, False
 End Select
End Sub

Private Sub Muestra2Objetos(b1 As Boolean, _
    b2 As Boolean, b3 As Boolean, b4 As Boolean, _
    b5 As Boolean, b6 As Boolean, b7 As Boolean, _
    b8 As Boolean, b9 As Boolean, b10 As Boolean, b11 As Boolean)
    
        Label1.Visible = b1
        Label2.Visible = b2
        Label3.Visible = b3
        Label4.Visible = b11
        Text1.Visible = b4
        Text2.Visible = b5
        Text3.Visible = b6
        cmdVer.Visible = b7
        PBar.Visible = b8
        Frame1.Visible = b9
        Frame2.Visible = b10

    End Sub




'Ejecuta el informe
'=======================================================
Private Sub cmdVer_Click()
'=======================================================

Screen.MousePointer = vbHourglass

Select Case mInfo
    Case kInfoDia
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        mInforme (kInfoDia)
        
    Case kInfoComercio
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        'Necesita el No de comercio
        If Not IsNumeric(Text3.Text) Then
            Text3.SetFocus
            GoTo sale
        End If
        mInforme (kInfoComercio)
    
    Case kInfoComercio2
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        'Necesita el No de comercio
        If Not IsNumeric(Text3.Text) Then
            Text3.SetFocus
            GoTo sale
        End If
        mInforme (kInfoComercio2)
    
    Case kInfoUnSocio
        mInformeSocio
    
    Case kInfoHistoriaUnSocio
        mInformeHistoriaUnSocio
    
    Case kInfoHistoriaUnSocio2
        mInformeHistoriaUnSocio2
    
    Case kInfoTodosSOcios
        mInformeTodosSocios (1)
    
    Case kInfoTodosSocios2
        mInformeTodosSocios (2)
    
    Case kInfoOrdAnuladas
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        mInfoOrdAnuladas
    
    Case kInfoCuentaSocios
        mInformeCuentaSocios
    
    Case kInfoCobros
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        Call mInfoCobros
    
    Case kInfoCobros2
        If Not VerificaFechas Then
            Text1.SetFocus
             GoTo sale
        End If
        mInfoCobros2
    
    Case kInfoCobros3
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        Call mInfoCobros3
        
    Case kInfoCobrosPorComercio
        'las fechas tienen que estar ok
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        
        'Necesita el No de comercio
        If Not IsNumeric(Text3.Text) Then
            Text3.SetFocus
            GoTo sale
        End If
 
        Call mInfoCobrosPorComercio
      
    Case kInfoRecargos
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        Call mInfoRecargos
    
    Case kInfoResumen
        mInfoResumen
    
    Case kInfoResumen2      'jefatura
        mInfoResumen2 (1)
    
    Case kInfoResumen3      'centro
        mInfoResumen2 (2)
    
    Case kInfoGastos
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        Call mInfoGastos(1)
    
    Case kInfoGastosAdm
        If Not VerificaFechas Then
            Text1.SetFocus
            GoTo sale
        End If
        Call mInfoGastos(2)
End Select
sale:
Screen.MousePointer = vbDefault

End Sub

'=======================================================
Private Function VerificaFechas() As Boolean
'=======================================================
If CDate(Text1.Text) > CDate(Text2.Text) Then
    VerificaFechas = False
Else
    VerificaFechas = True
End If

End Function


Private Sub CierraTodo()
    If adoClie.State = adStateOpen Then adoClie.Close
    Set adoClie = Nothing
    If adoCom.State = adStateOpen Then adoCom.Close
    Set adoCom = Nothing
    If adoC.State = adStateOpen Then adoC.Close
    Set adoC = Nothing
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
    If adoE.State = adStateOpen Then adoE.Close
    Set adoE = Nothing
    
    Set adoCmd = Nothing
    
    Set cOrd = Nothing
    Set cPag = Nothing
    Set cCom = Nothing
    Set cSocio = Nothing

End Sub


Private Sub Form_Unload(Cancel As Integer)
   CierraTodo
End Sub

'=======================================================
Private Sub Text1_change()
'=======================================================
    Select Case mInfo
        Case kInfoDia
            Call FormatoTexto(Text1)
        Case kInfoCobros, kInfoCobros2, kInfoCobros3, kInfoCobrosPorComercio, kInfoRecargos
            Call FormatoTexto(Text1)
        Case kInfoComercio, kInfoComercio2
            Call FormatoTexto(Text1)
        Case kInfoUnSocio
        Case kInfoTodosSOcios
        Case kInfoCuentaSocios
            Call FormatoTexto(Text1)
        Case kInfoResumen
    End Select
End Sub



Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub


'=======================================================
Private Sub Text2_Change()
'=======================================================
        Select Case mInfo
        Case kInfoDia
            Call FormatoTexto(Text2)
       Case kInfoCobros, kInfoCobros2, kInfoCobros3, kInfoCobrosPorComercio, kInfoRecargos
            Call FormatoTexto(Text2)
        Case kInfoComercio, kInfoComercio2
            Call FormatoTexto(Text2)
        Case kInfoUnSocio
        Case kInfoTodosSOcios
    End Select
End Sub


'=======================================================
Private Sub Text1_Validate(Cancel As Boolean)
'=======================================================
    Select Case mInfo
        Case kInfoDia
            If Not IsDate(Text1.Text) Then Cancel = True
        Case kInfoCuentaSocios
            If Not IsDate(Text1.Text) Then Cancel = True
       Case kInfoCobros, kInfoCobros2, kInfoCobros3, kInfoCobrosPorComercio, kInfoRecargos
            If Not IsDate(Text1.Text) Then Cancel = True
        Case kInfoComercio, kInfoComercio2
            If Not IsDate(Text1.Text) Then Cancel = True
             
        Case kInfoUnSocio
            If Len(Text1.Text) = 0 Then Text1.Text = 0
            If Not IsNumeric(Text1.Text) Then Cancel = True
        Case kInfoTodosSOcios
            If Len(Text1.Text) = 0 Then Text1.Text = 0
            If Not IsNumeric(Text1.Text) Then Cancel = True
            If Not (CLng(Text1.Text) = 0 Or Len(Text1.Text) = 6) Then Cancel = True
      Case kInfoResumen
            If Len(Text1.Text) = 0 Then Text1.Text = 0
            If Not IsNumeric(Text1.Text) Then Cancel = True
            If Not (CLng(Text1.Text) = 0 Or Len(Text1.Text) = 6) Then Cancel = True
   End Select


End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case mInfo
        Case kInfoUnSocio, kInfoHistoriaUnSocio
            If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrSocAlf6
                fjMuestraTabla.Show
           ElseIf KeyCode = 114 Then   'F3
                vpMuestraTabla = kMstrSocPorNC3    'Por No Cobro
                fjMuestraTabla.Show
            End If
    End Select
End Sub

'=======================================================
Private Sub Text2_Validate(Cancel As Boolean)
'=======================================================
    Select Case mInfo
        Case kInfoDia
            If Not IsDate(Text2.Text) Then Cancel = True
        Case kInfoCobros, kInfoCobros2, kInfoCobros3, kInfoCobrosPorComercio, kInfoRecargos
            If Not IsDate(Text2.Text) Then Cancel = True
        Case kInfoComercio, kInfoComercio2
            If Not IsDate(Text2.Text) Then Cancel = True
        Case kInfoUnSocio
        Case kInfoTodosSOcios
    End Select
End Sub



'=======================================================
Private Sub FormatoTexto(sText As TextBox)
'=======================================================
        If Len(sText.Text) = 2 Then
            sText.Text = Left(sText.Text, 2) & "/"
            sText.SelStart = 3
        ElseIf Len(sText.Text) = 5 Then
            sText.Text = Left(sText.Text, 5) & "/"
            sText.SelStart = 6
        End If

End Sub


'=======================================================
Private Sub Text3_Validate(Cancel As Boolean)
'=======================================================
    Select Case mInfo
        Case kInfoDia
        Case kInfoComercio, kInfoComercio2, kInfoCobrosPorComercio
            If Not IsNumeric(Text3.Text) Then Cancel = True
        Case kInfoUnSocio
        Case kInfoTodosSOcios
        Case kInfoCobros
    End Select
End Sub


Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case mInfo
        Case kInfoCobrosPorComercio, kInfoComercio, kInfoComercio2
            If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrComerc4
                fjMuestraTabla.Show
            End If
    End Select
End Sub






'=======================================================
Private Sub mInformeTodosSocios(nPrm As Byte)
'=======================================================
'nprm=1 Todos los socios: atrasado actual cuota ayuda total
'nrpm=2 Todos los socios 2: vales carnic cuota ayuda credito total

Dim n1 As Integer, n2 As Integer

Dim sfecha As Date
Dim nM As Integer
Dim sTitFecha
Dim sTitTipo

Dim bSinFechaFinal As Boolean

nM = Left(Text1.Text, 2)
If nM = 0 Then              'todos
    sTitFecha = ""
    bSinFechaFinal = True
    sfecha = Format(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac, "short date")
ElseIf nM > 12 Or nM < 1 Then
    Text1.SetFocus
    Exit Sub
Else
    sfecha = Format(vpnPrspHst & "/" & Left(Text1.Text, 2) & "/" & Right(Text1.Text, 4), "short date")
    sTitFecha = " hasta presup " & Left(Text1.Text, 2) & "/" & Right(Text1.Text, 4)
    bSinFechaFinal = False
End If

Label3.Caption = "7.Verificando..."
Label3.Refresh
cOrd.msInicia

'1) TOMA TODAS LAS ORDENES EN ADOORDENES
Label3.Caption = "6.Carga Ordenes..."
Label3.Refresh
Dim sMom As String
If cOrd.adoOrdenes.State = adStateOpen Then cOrd.adoOrdenes.Close
sMom = "SELECT *, cdate('01/02/1900') as Otro " & _
        "FROM TBL_Ordenes " & _
        "INNER JOIN tbl_socios ON " & _
        "tbl_socios.nrosoc = tbl_ordenes.ord_nrosoc " & _
        "WHERE year(ord_cerro) < 1901;"
cOrd.adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic



If cOrd.adoOrdenes.RecordCount = 0 Then
    MsgBox "No tiene Ordenes"
    cOrd.msTermina
    Exit Sub
End If
cOrd.msColocaFechaVtoEnOtro
Debug.Print " Ordenes en adoOrdenes " & cOrd.adoOrdenes.RecordCount

  'Set fjMome.DataGrid1.DataSource = cORD.adoOrdenes
  'fjMome.Show
  'Exit Sub


Label3.Caption = "5.Preparando ordenes..."
Label3.Refresh

'2) FILTRA LAS ORDENES POR CATEGORIA EN AdoOrdenes
For n1 = 0 To 5
    If Option1(n1) = True Then n2 = n1
Next n1
Select Case n2
    Case 0
        cOrd.adoOrdenes.Filter = "CodSitLab=1"
        sTitTipo = " Activos"
    Case 1
        cOrd.adoOrdenes.Filter = "CodSitLab=2"
        sTitTipo = " Comisiones"
    Case 2
        cOrd.adoOrdenes.Filter = "CodSitLab=4"
        sTitTipo = " Pensionistas"
    Case 3
        cOrd.adoOrdenes.Filter = "CodSitLab=3"
        sTitTipo = " Retirados"
    Case 4
        cOrd.adoOrdenes.Filter = "CodCatSoc=3"
        sTitTipo = " Cooperadores"
    Case Else
        sTitTipo = ""
End Select
Debug.Print "Filtro por categoria " & cOrd.adoOrdenes.RecordCount
  'Set fjMome.DataGrid1.DataSource = cORD.adoOrdenes
  'fjMome.Show
  'Exit Sub

'3) LAS PREPARA EN ADOm2
cOrd.msPreparaOrdenesAPagarEnAdoM2 (0)
Debug.Print "Cantidad registros en adom2 " & cOrd.adoM2.RecordCount

  'Set fjMome.DataGrid1.DataSource = cORD.adoM2
  'fjMome.Show
  'Exit Sub

'4) LAS COPIA PARA ADOC
Set adoC = cOrd.adoM2
cOrd.msTermina


'5) FILTRA LAS ORDENES SEGUN EL VENCIMIENTO EN AdoC
If Not bSinFechaFinal Then
    adoC.Filter = "VENCIM <= #" & sfecha & "#"
End If

'MsgBox adoC.RecordCount
Label3.Caption = "4.Crea Tabla..."
Label3.Refresh

'6) Crea una tabla nueva
Set adoD.ActiveConnection = adoConn
Set adoD = New ADODB.Recordset

If nPrm = 1 Then
        adoD.Fields.Append "Socio", adInteger, 2
        adoD.Fields.Append "Nombre", adChar, 30
        adoD.Fields.Append "Atrasado", adDouble
        adoD.Fields.Append "Actual", adDouble
        adoD.Fields.Append "Cuota", adDouble
        adoD.Fields.Append "Ayuda", adDouble
        adoD.Fields.Append "Total", adDouble
        adoD.Fields.Append "NCobro", adInteger, 2

Else
        adoD.Fields.Append "Socio", adInteger, 2
        adoD.Fields.Append "Nombre", adChar, 30
        adoD.Fields.Append "Credito", adDouble
        adoD.Fields.Append "Carnic", adDouble
        adoD.Fields.Append "Vales", adDouble
        adoD.Fields.Append "Cuota", adDouble
        adoD.Fields.Append "Ayuda", adDouble
        adoD.Fields.Append "Total", adDouble
        adoD.Fields.Append "NCobro", adInteger, 2

End If
 
 
adoD.CursorType = adOpenDynamic
adoD.LockType = adLockOptimistic
adoD.Open



Dim nSocio As Long
'para nprm=1
Dim dAtrasado As Double
Dim dActual As Double

Dim dCuota As Double
Dim dAyuda As Double
Dim sMome As Double

'para nprm=2
Dim dCredito As Double
Dim dCarnic As Double
Dim dVales As Double
    
adoC.Sort = "socio"
adoC.MoveFirst
nSocio = adoC!socio
dAtrasado = 0
dActual = 0
dCuota = 0
dAyuda = 0

dCredito = 0
dCarnic = 0
dVales = 0
Label3.Caption = "3.Calculando..."
Label3.Refresh

If nPrm = 1 Then
        Do While Not adoC.EOF
            If Not adoC!socio = nSocio Then
                'lograba
                adoD.AddNew
                adoD!socio = nSocio
                adoD!atrasado = dAtrasado
                adoD!actual = dActual
                adoD!cuota = dCuota
                adoD!ayuda = dAyuda
                adoD!Total = dAtrasado + dActual + dCuota + dAyuda
                adoD.Update
                'inicializa
                nSocio = adoC!socio
                dAtrasado = 0
                dActual = 0
                dCuota = 0
                dAyuda = 0
            End If
            sMome = adoC!valorp
            If adoC!NoOrden = 1 Then
                dCuota = dCuota + sMome
            ElseIf adoC!NoOrden = 2 Then
                dAyuda = dAyuda + sMome
            Else
                If adoC!vencim < sfecha Then
                    dAtrasado = dAtrasado + sMome
                Else
                    dActual = dActual + sMome
                End If
            End If
            adoC.MoveNext
        Loop
Else
        Do While Not adoC.EOF
            If Not adoC!socio = nSocio Then
                'lograba
                adoD.AddNew
                adoD!socio = nSocio
                adoD!credito = dCredito
                adoD!carnic = dCarnic
                adoD!vales = dVales
                adoD!cuota = dCuota
                adoD!ayuda = dAyuda
                adoD!Total = dCredito + dCarnic + dVales + dCuota + dAyuda
                adoD.Update
                'inicializa
                nSocio = adoC!socio
                dCredito = 0
                dCarnic = 0
                dVales = 0
                dCuota = 0
                dAyuda = 0
            End If
            sMome = adoC!valorp
            If adoC!NoOrden = 1 Then
                dCuota = dCuota + sMome
            ElseIf adoC!NoOrden = 2 Then
                dAyuda = dAyuda + sMome
            Else
                If adoC!comercio = 190 Then     'vales
                    dVales = dVales + sMome
                ElseIf adoC!comercio = 115 Then     'carnic
                    dCarnic = dCarnic + sMome
                Else        'credito
                    dCredito = dCredito + sMome
                End If
            End If
            adoC.MoveNext
        Loop
End If
adoC.Close
Set adoC = Nothing

Label3.Caption = "2.Nombre Socio..."
Label3.Refresh
adoClie.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
adoD.MoveFirst
Do While Not adoD.EOF
        adoClie.MoveFirst
        adoClie.Find "NroSoc =" & adoD!socio
        If Not adoClie.EOF Then
            adoD!nombre = Left(Trim(adoClie!Apellido) & "  " & Trim(adoClie!nombre), 30)
        Else
            adoD!nombre = "Desconocido"
        End If
        adoD!nCobro = adoClie!NroCob
adoD.MoveNext
Loop

Label3.Caption = "1.Cargando Reporte..."
Label3.Refresh
adoClie.Close
Set adoClie = Nothing

 ' Set fjMome.DataGrid1.DataSource = adoD
 ' fjMome.Show
 ' Exit Sub


'7) ORDENA
If Option2(0).Value Then
    adoD.Sort = "socio"
ElseIf Option2(1).Value Then
    adoD.Sort = "nombre"
ElseIf Option2(2).Value Then
    adoD.Sort = "NCobro"
End If

'8) MUESTRA EL REPORTE
If nPrm = 1 Then
        drTodosLosSocios.Caption = "Informe de Socios " & sTitTipo & sTitFecha & " al " & Date
        
        drTodosLosSocios.Title = "Informe de Socios " & sTitTipo & sTitFecha & " al " & Date
        
        Set drTodosLosSocios.DataSource = adoD
        drTodosLosSocios.DataMember = ""
        drTodosLosSocios.Sections(3).Controls(1).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(1).DataField = "socio"
        drTodosLosSocios.Sections(3).Controls(2).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(2).DataField = "nombre"
        drTodosLosSocios.Sections(3).Controls(3).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(3).DataField = "atrasado"
        drTodosLosSocios.Sections(3).Controls(4).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(4).DataField = "actual"
        drTodosLosSocios.Sections(3).Controls(5).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(5).DataField = "cuota"
        drTodosLosSocios.Sections(3).Controls(6).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(6).DataField = "ayuda"
        drTodosLosSocios.Sections(3).Controls(7).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(7).DataField = "total"
        drTodosLosSocios.Sections(3).Controls(8).DataMember = ""
        drTodosLosSocios.Sections(3).Controls(8).DataField = "NCobro"
        'totales
        drTodosLosSocios.Sections(5).Controls(1).DataMember = ""
        drTodosLosSocios.Sections(5).Controls(1).DataField = "atrasado"
        drTodosLosSocios.Sections(5).Controls(3).DataMember = ""
        drTodosLosSocios.Sections(5).Controls(3).DataField = "actual"
        drTodosLosSocios.Sections(5).Controls(4).DataMember = ""
        drTodosLosSocios.Sections(5).Controls(4).DataField = "cuota"
        drTodosLosSocios.Sections(5).Controls(5).DataMember = ""
        drTodosLosSocios.Sections(5).Controls(5).DataField = "ayuda"
        drTodosLosSocios.Sections(5).Controls(6).DataMember = ""
        drTodosLosSocios.Sections(5).Controls(6).DataField = "total"
        
        drTodosLosSocios.Refresh
        Screen.MousePointer = vbDefault
        
        drTodosLosSocios.Show
Else
        drTodosLosSocios2.Caption = "Informe de Socios " & sTitTipo & sTitFecha & " al " & Date
        
        drTodosLosSocios2.Title = "Informe de Socios " & sTitTipo & sTitFecha & " al " & Date
        
        Set drTodosLosSocios2.DataSource = adoD
        drTodosLosSocios2.DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(1).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(1).DataField = "socio"
        drTodosLosSocios2.Sections(3).Controls(2).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(2).DataField = "nombre"
        drTodosLosSocios2.Sections(3).Controls(3).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(3).DataField = "credito"
        drTodosLosSocios2.Sections(3).Controls(4).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(4).DataField = "carnic"
        drTodosLosSocios2.Sections(3).Controls(5).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(5).DataField = "vales"
        drTodosLosSocios2.Sections(3).Controls(6).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(6).DataField = "cuota"
        drTodosLosSocios2.Sections(3).Controls(7).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(7).DataField = "ayuda"
        drTodosLosSocios2.Sections(3).Controls(8).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(8).DataField = "NCobro"
        drTodosLosSocios2.Sections(3).Controls(9).DataMember = ""
        drTodosLosSocios2.Sections(3).Controls(9).DataField = "total"
        'totales
        drTodosLosSocios2.Sections(5).Controls(1).DataMember = ""
        drTodosLosSocios2.Sections(5).Controls(1).DataField = "credito"
        drTodosLosSocios2.Sections(5).Controls(3).DataMember = ""
        drTodosLosSocios2.Sections(5).Controls(3).DataField = "carnic"
        drTodosLosSocios2.Sections(5).Controls(4).DataMember = ""
        drTodosLosSocios2.Sections(5).Controls(4).DataField = "vales"
        drTodosLosSocios2.Sections(5).Controls(5).DataMember = ""
        drTodosLosSocios2.Sections(5).Controls(5).DataField = "cuota"
        drTodosLosSocios2.Sections(5).Controls(6).DataMember = ""
        drTodosLosSocios2.Sections(5).Controls(6).DataField = "ayuda"
        drTodosLosSocios2.Sections(5).Controls(7).DataMember = ""
        drTodosLosSocios2.Sections(5).Controls(7).DataField = "total"
        
        drTodosLosSocios2.Refresh
        Screen.MousePointer = vbDefault
        
        drTodosLosSocios2.Show

End If
Label3.Caption = ""
Label3.Refresh
'Set DataGrid1.DataSource = adoD
Exit Sub


End Sub





'=======================================================
Private Sub mInformeSocio()
'=======================================================
Dim sNombre As String
            Screen.MousePointer = vbHourglass

    adoClie.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    adoCom.Open "select * FROM tbl_Comercios ORDER BY codigo;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
 '6.1) Coloca el nombre
        adoClie.MoveFirst
        adoClie.Find "NroSoc =" & Text1.Text
        If Not adoClie.EOF Then
            sNombre = Left(Trim(adoClie!Apellido) & "  " & _
                Trim(adoClie!nombre), 30) & "  Nro.Cob. " & adoClie!NroCob
        Else
            sNombre = "Desconocido"
        End If
        adoClie.Close
        Set adoClie = Nothing
 '1)BUSCA LAS ORDENES QUE TIENE EL SOCIO
    cOrd.msInicia
    
    cOrd.vlNroSoc = Text1.Text
    'If cOrd.vlNroSoc = 3451 Then
    '    Debug.Print "hola"
    'End If
    If Not cOrd.fBuscaOrdenesUnSocio Then     'HUBO PROBLEMAS
        MsgBox "2348s: Problemas al Buscar Ordenes"
        GoTo miFinal
    End If
    
    If cOrd.adoOrdenes.RecordCount = 0 Then
        MsgBox "No tiene Ordenes"
        GoTo miFinal
    End If
    
        'ojo momentan
    'Set DataGrid1.DataSource = cORD.adoOrdenes
    'Exit Sub
    
    '2.Prepara el ado
    cOrd.msPreparaOrdenesAPagarEnAdoM2 (5)
    Set adoC = cOrd.adoM2
    'Set DataGrid1.DataSource = adoC
    'Exit Sub
    
    '3) crea la tabla virtual
    Set adoD.ActiveConnection = adoConn
    Set adoD = New ADODB.Recordset
    Dim nM As Integer
    Dim mNomb As String
    Dim mTipo
    Dim mTam As Long
    
    '3.1) con los mismos campos
    For nM = 0 To adoC.Fields.Count - 1
        mNomb = adoC.Fields(nM).Name
        mTipo = adoC.Fields(nM).Type
        mTam = adoC.Fields(nM).DefinedSize
        adoD.Fields.Append mNomb, mTipo, mTam
    Next
    '3.2) con campos nuevos
    adoD.Fields.Append "nComercio", adChar, 30
    adoD.Fields.Append "Tipo", adChar, 30
    adoD.Fields.Append "Mensaje", adChar, 30
    adoD.Fields.Append "STot", adChar, 15
    adoD.Fields.Append "sFecha", adChar, 10
     
    adoD.CursorType = adOpenDynamic
    adoD.LockType = adLockOptimistic
    adoD.Open
    
    '4) LLena los campos
     
    adoC.MoveFirst
    mTam = adoC.Fields.Count - 1
    Do While Not adoC.EOF
            adoD.AddNew
            
            'los campos comunes
            For nM = 0 To mTam
                adoD(nM) = adoC(nM)
            Next
            'los campos especiales
            'tipo
            Select Case adoC("noorden")
                Case 0
                    adoD("Tipo") = "Orden"
                Case 1
                    adoD("Tipo") = "Cuota"
                Case 2
                   adoD("Tipo") = "Ayuda"
                Case 3
                   adoD("Tipo") = "Recargo"
                Case Else
                   adoD("Tipo") = ""
            End Select
            '6.4) Coloca el Comercio
            adoCom.MoveFirst
            adoCom.Find "Codigo =" & adoC!comercio
            If Not adoCom.EOF Then
                adoD!nComercio = "" & Left(Trim(adoCom!NombCom), 30)
            Else
                adoD!nComercio = ""
            End If
             adoD("Mensaje") = Trim(adoD("tipo") & adoD("ncomercio"))
            If Trim(adoD("moned")) = "" Then adoD("moned") = "$"
            adoC.MoveNext
    Loop
    
   'Coloca subtotales por dia de vencimiento
  Dim mSTot As Double
  Dim sfecha As String
  mSTot = 0
  adoD.MoveFirst
  sfecha = adoD!vencim
  adoD!sfecha = CStr(adoD!vencim)
  
  Do While Not adoD.EOF
        'cambio de fecha de vencimiento
        If Not adoD!vencim = sfecha Then
            adoD.MovePrevious
            adoD!sTot = Format(mSTot, "#,#0.00")
            adoD.MoveNext
            mSTot = 0
            sfecha = adoD!vencim
            adoD!sfecha = CStr(adoD!vencim)
        End If
        mSTot = mSTot + adoD!valorp
        adoD.MoveNext
    Loop
adoD.MovePrevious
adoD!sTot = Format(mSTot, "#,#0.00")

    adoD.Sort = "vencim"
    'Set DataGrid1.DataSource = adoC
  drInfoSocio.Caption = "Informe de  " & sNombre & "   Nro:" & Text1.Text

  drInfoSocio.Title = "Informe de:  " & sNombre & vbCrLf & " Socio Nro: " & Text1.Text & "   al " & Date

  Set drInfoSocio.DataSource = adoD
  drInfoSocio.DataMember = ""

  drInfoSocio.Sections(3).Controls(2).DataMember = ""
  drInfoSocio.Sections(3).Controls(2).DataField = "sfecha"
  drInfoSocio.Sections(3).Controls(1).DataMember = ""
  drInfoSocio.Sections(3).Controls(1).DataField = "noorden"
  drInfoSocio.Sections(3).Controls(3).DataMember = ""
  drInfoSocio.Sections(3).Controls(3).DataField = "nodepend"
  drInfoSocio.Sections(3).Controls(4).DataMember = ""
  drInfoSocio.Sections(3).Controls(4).DataField = "mensaje"
  drInfoSocio.Sections(3).Controls(5).DataMember = ""
  drInfoSocio.Sections(3).Controls(5).DataField = "valorp"
  drInfoSocio.Sections(3).Controls(6).DataMember = ""
  drInfoSocio.Sections(3).Controls(6).DataField = "valorme"
  drInfoSocio.Sections(3).Controls(7).DataMember = ""
  drInfoSocio.Sections(3).Controls(7).DataField = "moned"
  drInfoSocio.Sections(3).Controls(8).DataMember = ""
  drInfoSocio.Sections(3).Controls(8).DataField = "stot"
  drInfoSocio.Sections(3).Controls(9).DataMember = ""
  drInfoSocio.Sections(3).Controls(9).DataField = "cuota"
  'totales
  drInfoSocio.Sections(5).Controls(1).DataMember = ""
  drInfoSocio.Sections(5).Controls(1).DataField = "valorp"
  
  drInfoSocio.Refresh
        Screen.MousePointer = vbDefault
  
  drInfoSocio.Show
miFinal:
    cOrd.msTermina
    adoCom.Close
    Set adoCom = Nothing
    If adoC.State = adStateOpen Then adoC.Close
    Set adoC = Nothing
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
End Sub

'=======================================================
Private Sub mInforme(nPrm As Byte)
'=======================================================
    ' Listado de Comercio Ordenado por Fecha y No Orden
    ' Listado de Comercio Ordenado por Socio (con subtotal por socio)
    ' Listado de Ordenes emitidas en el dia.
    ' En todos los listados, lista todas las ordenes inclusive las pagadas.
    ' En las de Comercio NO lista las anuladas.
    
    Dim nM As Integer
    Dim sM As String
    Dim sM1 As String
    Dim sM2 As String
    Screen.MousePointer = vbHourglass

    '1) Elimina los registros de tbl_Inf01
    Dim adoCmd As New ADODB.Command
    adoCmd.CommandText = "delete * from tbl_info01"
    Set adoCmd.ActiveConnection = adoConn
    adoCmd.Execute
    Set adoCmd = Nothing
    
    '2) Copia los registros del dia
    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)
    If adoD.State = adStateOpen Then adoD.Close
    If nPrm = kInfoComercio Or nPrm = kInfoComercio2 Then
        Dim lCom As Long
        Dim tCom As String
        lCom = CLng(Text3.Text)
        tCom = cCom.BuscaComercio(lCom)
        Set cCom = Nothing
        sM = "select * FROM tbl_Ordenes" & _
            " WHERE ord_NroCom =" & CStr(lCom) & " AND" & _
            " ord_NroOrden > 2 AND" & _
            " NOT ord_tipo = 4 AND" & _
            " ord_FEmis BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
            'ORD_TIPO =4 ANULADA
            '13/5/3 elimino:          " ord_Cerro < #01/01/1901# AND" & _

        adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    ElseIf nPrm = kInfoDia Then
        sM = "select * FROM tbl_Ordenes" & _
            " WHERE ORD_FEmis BETWEEN #" & sM1 & "# AND #" & sM2 & "#" & _
            " AND ord_NroOrden > 2;"
            '13/5/03 elimino:  " AND ord_Cerro < #01/01/1901#;"
                    
        adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    
    '3) Si no hay registros termina
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        Exit Sub
    End If
    
    '31) Coloca en un vector la tabla situacion laboral
    If adoE.State = adStateOpen Then adoE.Close
    adoE.Open "select * FROM sLaboral;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    Dim kLab As Long
    kLab = adoE.RecordCount + 1
    Dim sLaboral() As String
    ReDim sLaboral(kLab)
    adoE.MoveFirst
    Do While Not adoE.EOF
        sLaboral(adoE!idsitlab) = adoE!Desc
        adoE.MoveNext
    Loop
    adoE.Close
    
 
    '4) Guarda los registros en la tabla tbl_Inf01
    If adoC.State = adStateOpen Then adoC.Close
    adoC.Open "select * FROM tbl_info01;", adoConn, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoD.MoveFirst
    Do While Not adoD.EOF
        adoC.AddNew
        For nM = 0 To adoD.Fields.Count - 1
            adoC(nM) = adoD(nM)
        Next
        adoC.Update
        adoD.MoveNext
    Loop
    adoD.Close
    
    '5) Prepara la barra de progreso
    PBar.Min = 0
    PBar.Max = adoC.RecordCount * 2 + 5
    PBar.Value = 5
    nM = 5
    
    '6) ARREGLA EL ADO
    If adoD.State = adStateOpen Then adoD.Close
    adoD.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    If adoE.State = adStateOpen Then adoE.Close
    adoE.Open "select * FROM tbl_Comercios ORDER BY codigo;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    adoC.MoveFirst
    'Dim sMome As String
    Do While Not adoC.EOF
        '6.1) Coloca el nombre
        adoD.MoveFirst
        adoD.Find "NroSoc =" & adoC!ord_nrosoc
        If Not adoD.EOF Then
            adoC!i_sNombre = Left(Trim(adoD!Apellido) & "  " & Trim(adoD!nombre), 30)
            adoC!i_sm2 = sLaboral(adoD!CodSitLab)   'situacion laboral
        Else
            adoC!i_sNombre = "Desconocido"
            adoC!i_sm2 = ""
        End If
        '6.2) Coloca los totales
        adoC!ord_cuota = adoC!ord_cuota * adoC!ORD_PLAN
        adoC!ord_mecuota = adoC!ord_mecuota * adoC!ORD_PLAN
        
        '6.3) Coloca las monedas
        Select Case adoC!ord_Mon
            Case "P"
                adoC!i_sMoneda = "Pesos"
                adoC!ord_mecuota = 0
                adoC!ORD_MEPagos = 0
            Case "D"
                adoC!i_sMoneda = "Dólares"
            Case "R"
                adoC!i_sMoneda = "Reales"
            Case "A"
                adoC!i_sMoneda = "Australes"
            Case "U"
                adoC!i_sMoneda = "U.Reajustables"
        End Select
        
        If nPrm = kInfoDia Then
                '6.4) Coloca el Comercio
                adoE.MoveFirst
                adoE.Find "Codigo =" & adoC!ORD_NroCom
                If Not adoE.EOF Then
                    adoC!i_sM1 = Left(Trim(adoE!NombCom), 30)
                Else
                    adoC!i_sM1 = "Desconocido"
                End If
                
                '6.4 Coloca la situacion laboral

        End If
        
        '6.5) completa el bucle
        adoC.MoveNext
        PBar.Value = nM
        nM = nM + 1
    Loop
    
    '7) Actualiza el adoC: la tabla tbl_Info01
    adoC.UpdateBatch
    adoC.MoveLast
    adoC.MoveFirst
    adoC.Close
    adoD.Close
    
    PBar.Max = 100
    PBar.Value = 100
    
    'INFORME DE CARNICERIA
    If nPrm = kInfoComercio Then
                 '8) Carga la tabla tbl_Inf01 en el ADO
                 If adoC.State = adStateOpen Then adoC.Close
                 adoC.Open "select * FROM tbl_info01 ORDER BY ord_Femis, ord_NroOrden;", adoConn, adOpenKeyset, adLockBatchOptimistic, adCmdText
                 adoC.MoveLast
                 adoC.MoveFirst
                 DoEvents
                 
                 
                 
                 '9) Muestra el informe de datos
                                     'un solo dia
                    If CDate(Text1.Text) = CDate(Text2.Text) Then
                        drInfCarnicDia.Caption = "Resumen de " & tCom & " del día " & Text1.Text
                        drInfCarnicDia.Title = "Resumen de  " & tCom & vbCrLf & "del día: " & Text1.Text
                     'varios dias
                    Else
                        drInfCarnicDia.Caption = "Resumen de " & tCom & "  desde: " & Text1.Text & "  hasta: " & Text2.Text
                        drInfCarnicDia.Title = "Resumen de " & tCom & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
                    End If
                 Set drInfCarnicDia.DataSource = adoC
                
                 drInfCarnicDia.Sections(3).Controls(1).DataField = adoC(2).Name
                 drInfCarnicDia.Sections(3).Controls(2).DataField = adoC(0).Name
                 drInfCarnicDia.Sections(3).Controls(3).DataField = adoC(4).Name
                 drInfCarnicDia.Sections(3).Controls(4).DataField = adoC(19).Name
                 drInfCarnicDia.Sections(3).Controls(5).DataField = "i_sM1"
                 drInfCarnicDia.Sections(3).Controls(6).DataField = "ord_FEmis"
                 drInfCarnicDia.Sections(5).Controls(1).DataField = adoC(4).Name
                 drInfCarnicDia.Refresh
                 Screen.MousePointer = vbDefault
                 DoEvents
                 drInfCarnicDia.Show
       ElseIf nPrm = kInfoComercio2 Then
                    '8) Carga la tabla tbl_Inf01 en el ADO
                 If adoC.State = adStateOpen Then adoC.Close
                 adoC.Open "select * FROM tbl_info01 ORDER BY ord_NroSoc,ord_Femis, ord_NroOrden;", adoConn, adOpenKeyset, adLockBatchOptimistic, adCmdText
                 adoC.MoveLast
                 adoC.MoveFirst
                 DoEvents
                'Coloca subtotales por socio
                  Dim mSTot As Double
                  Dim lSocio As Long
                  mSTot = 0
                  adoC.MoveFirst
                  lSocio = adoC!ord_nrosoc
                  
                  Do While Not adoC.EOF
                        'cambio de socio
                        If Not adoC!ord_nrosoc = lSocio Then
                            adoC.MovePrevious
                            adoC!i_sm2 = Format(mSTot, "#,#0.00")
                            adoC.MoveNext
                            mSTot = 0
                            lSocio = adoC!ord_nrosoc
                        End If
                        mSTot = mSTot + adoC!ord_cuota
                        adoC.MoveNext
                    Loop
                adoC.MovePrevious
                adoC!i_sm2 = Format(mSTot, "#,#0.00")

                  '9) Muestra el informe de datos
                                     'un solo dia
                    If CDate(Text1.Text) = CDate(Text2.Text) Then
                        drInfCarnicDia.Caption = "Resumen de " & tCom & " del día " & Text1.Text
                        drInfCarnicDia.Title = "Resumen de  " & tCom & vbCrLf & "del día: " & Text1.Text
                     'varios dias
                    Else
                        drInfCarnicDia.Caption = "Resumen de " & tCom & "  desde: " & Text1.Text & "  hasta: " & Text2.Text
                        drInfCarnicDia.Title = "Resumen de " & tCom & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
                    End If
                 Set drInfCarnicDia.DataSource = adoC
                
                 drInfCarnicDia.Sections(3).Controls(1).DataField = adoC(2).Name
                 drInfCarnicDia.Sections(3).Controls(2).DataField = adoC(0).Name
                 drInfCarnicDia.Sections(3).Controls(3).DataField = adoC(4).Name
                 drInfCarnicDia.Sections(3).Controls(4).DataField = adoC(19).Name
                 drInfCarnicDia.Sections(3).Controls(5).DataField = "i_sM2"
                 drInfCarnicDia.Sections(3).Controls(6).DataField = "ord_FEmis"
                 drInfCarnicDia.Sections(5).Controls(1).DataField = adoC(4).Name
                 drInfCarnicDia.Refresh
                 Screen.MousePointer = vbDefault
                 DoEvents
                 drInfCarnicDia.Show

  
    'INFORME DIARIO
    ElseIf nPrm = kInfoDia Then
                  Dim cn As New ADODB.Connection
                    Set cn = New ADODB.Connection
                    cn.CursorLocation = adUseClient
                    cn.Provider = "MSDATASHAPE"
                    cn.Open "dsn=jimmy"
        
        
                    With adoCmd
                        .ActiveConnection = cn
                        .CommandType = adCmdText
                        .CommandText = "SHAPE {select * from tbl_info01 ORDER BY i_sMoneda, ord_FEmis, ord_NroOrden;}  AS tbl_Info01 COMPUTE tbl_Info01 BY 'i_sMoneda'"
                        .Execute
                    End With
               
                    With adoC
                        .ActiveConnection = cn
                        .CursorLocation = adUseClient
                        .Open adoCmd
                    End With
                    Set adoD = adoC(0).Value
                    'MsgBox adoC(0).Name & vbCrLf & adoC(1).Name & vbCrLf & adoC(0)(0).Name
                     
                    '9) Muestra el informe de datos
                    drInfDia2.Hide
                    'un solo dia
                    If CDate(Text1.Text) = CDate(Text2.Text) Then
                        drInfDia2.Caption = "Resumen de Movimentos del día: " & Text1.Text
                        drInfDia2.Title = "Resumen de Movimientos " & vbCrLf & "del día: " & Text1.Text
                     'varios dias
                    Else
                        drInfDia2.Caption = "Resumen de Movimentos desde: " & Text1.Text & "  hasta: " & Text2.Text
                        drInfDia2.Title = "Resumen de Movimentos" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
                    End If
                    Set drInfDia2.DataSource = adoC
                    drInfDia2.DataMember = ""
                  
                    'GRUPO
                    'drInfDia2.Sections(3).Controls(1).DataMember = ""
                    'drInfDia2.Sections(3).Controls(1).DataField = "i_smoneda"
                    'DETALLES
                    drInfDia2.Sections(4).Controls(1).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(1).DataField = "ord_FEmis"
                    drInfDia2.Sections(4).Controls(2).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(2).DataField = "ord_NroOrden"
                    drInfDia2.Sections(4).Controls(3).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(3).DataField = "ord_NroSoc"
                    drInfDia2.Sections(4).Controls(4).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(4).DataField = "i_sNombre"
                    drInfDia2.Sections(4).Controls(5).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(5).DataField = "Ord_Cuota"
                    drInfDia2.Sections(4).Controls(6).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(6).DataField = "Ord_meCuota"
                    drInfDia2.Sections(4).Controls(7).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(7).DataField = "i_sM1"           'COMERCIO
                    drInfDia2.Sections(4).Controls(8).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(8).DataField = "ord_plan"           'COMERCIO
                    drInfDia2.Sections(4).Controls(9).DataMember = "tbl_Info01"
                    drInfDia2.Sections(4).Controls(9).DataField = "i_sm2"           'COMERCIO
                     'PIE DE INFORME
                    drInfDia2.Sections(7).Controls(2).DataMember = "tbl_Info01"
                    drInfDia2.Sections(7).Controls(2).DataField = "ord_cuota"
                    drInfDia2.Sections(7).Controls(3).DataMember = "tbl_Info01"
                    drInfDia2.Sections(7).Controls(3).DataField = "ord_meCuota"
    
                    drInfDia2.Refresh
                    Screen.MousePointer = vbDefault
                    DoEvents
                    drInfDia2.Show
End If
End Sub


'=======================================================
Private Sub EliminaTablaInf01()
'=======================================================
        Dim mCatalog As New ADOX.Catalog
        Dim bExiste As Boolean
        Dim nP As Integer
        Set mCatalog.ActiveConnection = adoConn
    
        For nP = 0 To mCatalog.Tables.Count - 1
                If mCatalog.Tables(nP).Name = "tbl_inf01" Then
                    bExiste = True
                    Exit For
                End If
        Next nP
           
     If bExiste Then
            Dim adoBorrar As New ADODB.Command
            adoBorrar.CommandText = "drop table tbl_info01"
            Set adoBorrar.ActiveConnection = adoConn
            adoBorrar.Execute
        End If
        Set adoBorrar = Nothing
        Set mCatalog = Nothing

End Sub


'=======================================================
Private Sub mInfoOrdAnuladas()
'=======================================================
    Dim sM1 As String
    Dim sM2 As String
    
    Screen.MousePointer = vbHourglass
    
    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String
    'se puede cambiar para ord_tipo=4
    sM = "select * FROM tbl_Pagos INNER JOIN tbl_Socios" & _
            " ON tbl_Socios.NroSoc = tbl_Pagos.pag_NroSoc" & _
            " WHERE pag_Motivo = 4 AND" & _
            " pag_Fecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
    adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
  
    adoD.Sort = "pag_Fecha, pag_NroOrden"
    
    ' Muestra el informe de datos
    'un solo dia
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfOrdAnul.Caption = "Resumen de Ordenes Anuladas el día: " & Text1.Text
        drInfOrdAnul.Title = "Resumen de Ordenes Anuladas " & vbCrLf & "del día: " & Text1.Text
     'varios dias
    Else
        drInfOrdAnul.Caption = "Resumen de Ordenes Anuladas desde: " & Text1.Text & "  hasta: " & Text2.Text
        drInfOrdAnul.Title = "Resumen de Ordenes Anuladas" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
    End If
    
    Set drInfOrdAnul.DataSource = adoD
    drInfOrdAnul.DataMember = ""
    
    drInfOrdAnul.Sections(3).Controls(1).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(1).DataField = "pag_Fecha"
    drInfOrdAnul.Sections(3).Controls(2).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(2).DataField = "pag_NroORden"
    drInfOrdAnul.Sections(3).Controls(3).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(3).DataField = "pag_NroSoc"
    drInfOrdAnul.Sections(3).Controls(4).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(4).DataField = "Apellido"
    drInfOrdAnul.Sections(3).Controls(5).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(5).DataField = "pag_valor"
    drInfOrdAnul.Sections(3).Controls(6).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(6).DataField = "pag_valME"
    drInfOrdAnul.Sections(3).Controls(7).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(7).DataField = "pag_NroCom"
    drInfOrdAnul.Sections(3).Controls(8).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(8).DataField = "Nombre"
    drInfOrdAnul.Sections(3).Controls(9).DataMember = ""
    drInfOrdAnul.Sections(3).Controls(9).DataField = "pag_det"
    'totales
    drInfOrdAnul.Sections(5).Controls(1).DataMember = ""
    drInfOrdAnul.Sections(5).Controls(1).DataField = "pag_valor"
    drInfOrdAnul.Sections(5).Controls(2).DataMember = ""
    drInfOrdAnul.Sections(5).Controls(2).DataField = "pag_valME"
    
    drInfOrdAnul.Refresh
    Screen.MousePointer = vbDefault
    
    drInfOrdAnul.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
End Sub



'=======================================================
Private Sub mInfoCobros()
'=======================================================
    Dim sM1 As String
    Dim sM2 As String
    
    Screen.MousePointer = vbHourglass
    PBar.Visible = False
    Label5.Caption = "Espere..."
    Label5.Refresh

    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String
    sM = "select *,space(25) as apellido, space(25) as nombre FROM tbl_Pagos" & _
            " WHERE not pag_Motivo = 4 AND" & _
            " NOT pag_Motivo = 7 AND" & _
            " pag_Fecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
    adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
    
    'coloca los nombres
    Label5.Caption = "Nombres..."
    Label5.Refresh
    cSocio.mfAbreTablaSociosOrdenSocio
    adoD.MoveFirst
    Do While Not adoD.EOF
        cSocio.vlNroSoc = adoD!pag_NroSoc
        If Not cSocio.mfBuscaSocio Then
            adoD!Apellido = "NO ENCONTRADO"
            adoD!nombre = ""
        Else
            adoD!Apellido = cSocio.vsApellido
            adoD!nombre = cSocio.vsNombre
            adoD.Update
        End If
        adoD.MoveNext
    Loop
    
    
    adoD.Sort = "pag_Fecha, pag_NroOrden"
    'adoD.Filter = "pag_motivo <> 7"     'sin recargo

   
    ' Muestra el informe de datos
    'un solo dia
    Label5.Caption = "Armando..."
    Label5.Refresh
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfCobros.Caption = "Resumen de Cobros del día: " & Text1.Text
        drInfCobros.Title = "Resumen de Cobros " & vbCrLf & "del día: " & Text1.Text
     'varios dias
    Else
        drInfCobros.Caption = "Resumen de Cobros desde: " & Text1.Text & "  hasta: " & Text2.Text
        drInfCobros.Title = "Resumen de Cobros" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
    End If
    
    Set drInfCobros.DataSource = adoD
    drInfCobros.DataMember = ""
    
    drInfCobros.Sections(3).Controls(1).DataMember = ""
    drInfCobros.Sections(3).Controls(1).DataField = "pag_Fecha"
    drInfCobros.Sections(3).Controls(2).DataMember = ""
    drInfCobros.Sections(3).Controls(2).DataField = "pag_NroORden"
    drInfCobros.Sections(3).Controls(3).DataMember = ""
    drInfCobros.Sections(3).Controls(3).DataField = "pag_NroSoc"
    drInfCobros.Sections(3).Controls(4).DataMember = ""
    drInfCobros.Sections(3).Controls(4).DataField = "Apellido"
    drInfCobros.Sections(3).Controls(5).DataMember = ""
    drInfCobros.Sections(3).Controls(5).DataField = "pag_valor"
    drInfCobros.Sections(3).Controls(6).DataMember = ""
    drInfCobros.Sections(3).Controls(6).DataField = "pag_valME"
    drInfCobros.Sections(3).Controls(7).DataMember = ""
    drInfCobros.Sections(3).Controls(7).DataField = "pag_NroCom"
    drInfCobros.Sections(3).Controls(8).DataMember = ""
    drInfCobros.Sections(3).Controls(8).DataField = "NOMBRE"
    drInfCobros.Sections(3).Controls(9).DataMember = ""
    drInfCobros.Sections(3).Controls(9).DataField = "pag_det"
    drInfCobros.Sections(3).Controls(10).DataMember = ""
    drInfCobros.Sections(3).Controls(10).DataField = "pag_NroPago"
   'totales
    drInfCobros.Sections(5).Controls(1).DataMember = ""
    drInfCobros.Sections(5).Controls(1).DataField = "pag_valor"
    drInfCobros.Sections(5).Controls(2).DataMember = ""
    drInfCobros.Sections(5).Controls(2).DataField = "pag_valME"
    
    drInfCobros.Refresh
    Screen.MousePointer = vbDefault
    Label5.Caption = ""
    Label5.Refresh

    drInfCobros.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
    Unload Me
End Sub



'=======================================================
Private Sub mInformeHistoriaUnSocio()
'=======================================================
    Dim sNombre As String
    Screen.MousePointer = vbHourglass
    
    If adoClie.State = adStateOpen Then adoClie.Close
    adoClie.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    
    If adoCom.State = adStateOpen Then adoCom.Close
    adoCom.Open "select * FROM tbl_Comercios ORDER BY codigo;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    '===========================================================
    '1) Coloca el nombre del socio pedido.
    adoClie.MoveFirst
    adoClie.Find "NroSoc =" & Text1.Text
    If Not adoClie.EOF Then
        sNombre = Left(Trim(adoClie!Apellido) & "  " & _
        Trim(adoClie!nombre), 30) & "  Nro.Cob. " & adoClie!NroCob
    Else
        sNombre = "Desconocido"
    End If
    adoClie.Close
    Set adoClie = Nothing
 
 
    '===========================================================
    '2)BUSCA LAS ORDENES QUE TIENE EL SOCIO
    cOrd.msInicia
    cOrd.vlNroSoc = Text1.Text
    
    If Not cOrd.fBuscaTodasLasOrdenesUnSocio Then      'HUBO PROBLEMAS
        MsgBox "2328s(2): Problemas al Buscar Ordenes  "
        GoTo miFinal
    End If
    
    'If cOrd.adoOrdenes.RecordCount = 0 Then
    '    MsgBox "No tiene Ordenes"
    '    GoTo miPagos
    'End If
    
    'ojo momentan
    'Set DataGrid1.DataSource = cORD.adoOrdenes
    'Exit Sub
    '2.Prepara el ado
    
    'Ordena las ordenes por el total inicial, sin recargo o ent. ctas.
    'con fecha de emisión
    cOrd.msPreparaOrdenesEnAdoM2
    If cOrd.adoM2.RecordCount < 1 Then
        Exit Sub
    End If
    
    '===========================================================
    '3) Las copia en adoC
    Set adoC = cOrd.adoM2
    
    'Set fjMome.DataGrid1.DataSource = adoC     'cOrd.adoOrdenes
    'fjMome.Show
    'Exit Sub
    
    '===========================================================
    '4) crea la tabla virtual adoD
    Set adoD.ActiveConnection = adoConn
    Set adoD = New ADODB.Recordset
    Dim nM As Integer
    Dim mNomb As String
    Dim mTipo
    Dim mTam As Long
    '4.1) con los mismos campos
    For nM = 0 To adoC.Fields.Count - 1
        mNomb = adoC.Fields(nM).Name
        mTipo = adoC.Fields(nM).Type
        mTam = adoC.Fields(nM).DefinedSize
        adoD.Fields.Append mNomb, mTipo, mTam
    Next
    '4.2) con campos nuevos
    adoD.Fields.Append "nComercio", adChar, 30
    adoD.Fields.Append "Tipo", adChar, 30
    adoD.Fields.Append "Mensaje", adChar, 30
    adoD.Fields.Append "Haber", adChar, 30              'es un char
    adoD.Fields.Append "sFecha", adChar, 10
    adoD.Fields.Append "Debe", adChar, 30               'es un char
    adoD.Fields.Append "fDebe", adSingle
    adoD.Fields.Append "fHaber", adSingle
    adoD.Fields.Append "fDife", adSingle
     
    adoD.CursorType = adOpenDynamic
    adoD.LockType = adLockOptimistic
    adoD.Open
    '4.3) LLena los campos
     
    adoC.MoveFirst
    mTam = adoC.Fields.Count - 1
    Do While Not adoC.EOF
        adoD.AddNew
        'los campos comunes
        For nM = 0 To mTam
            adoD(nM) = adoC(nM)
        Next
        'los campos especiales
        'tipo
        Select Case adoC("noorden")
            Case 0
                adoD("Tipo") = "Orden"
            Case 1
                adoD("Tipo") = "Cuota"
            Case 2
               adoD("Tipo") = "Ayuda"
            Case 3  'No deberian aparecer de estas. Recargos
               adoD("Tipo") = "???????"
            Case Else
               adoD("Tipo") = "Orden"
        End Select
        '4.4) Coloca el Comercio
        adoCom.MoveFirst
        adoCom.Find "Codigo =" & adoC!comercio
        If Not adoCom.EOF Then
            adoD!nComercio = "" & Left(Trim(adoCom!NombCom), 30)
        Else
            adoD!nComercio = ""
        End If
        adoD!valorp = adoD!valorp + adoD!valorME
        adoD!haber = ""
        adoD!debe = Format(adoD!valorp, "#,#0.00")
        adoD!fdebe = adoD!valorp
        adoD!fdife = adoD!valorp
     adoC.MoveNext
    Loop
    adoD.Update
    adoC.Close
    'Set fjMome.DataGrid1.DataSource = adoD     'cOrd.adoOrdenes
    'fjMome.Show
    'Exit Sub
  

    '===========================================================
    '5) Toma los datos de pagos
    cPag.mfAbrePagosDeUnSocio (CStr(CLng(Text1.Text)))
    'If cPag.adoPagos.RecordCount < 1 Then
    '    MsgBox "No tiene Pagos"
    '    GoTo miFinal
    'End If
    
    '6) Los pasa al adoc
    Set adoC = cPag.adoPagos
    adoC.Sort = "pag_auto"
    If adoC.RecordCount > 0 Then
        adoC.MoveFirst
        '7) Los agrega en adod
        Do While Not adoC.EOF
            adoD.AddNew
            adoD!NoOrden = adoC!pag_NroOrden
            adoD!moned = adoC!pag_mon
            adoD!vencim = adoC!pag_Fecha
            Select Case adoC!pag_Motivo
                Case 4
                    adoD!tipo = "Anulac Orden"
                Case 5
                    If adoD!NoOrden < 3 Then
                        adoD!tipo = "Pago Cuota"
                    Else
                      adoD!tipo = "Pago Orden"
                    End If
                Case 6
                    adoD!tipo = "Pago E.Cta."
                Case 8
                    adoD!tipo = "Anula Pago "
                Case 7
                    adoD!tipo = adoC!pag_det
                Case Else
                    adoD!tipo = "Pago ?"
            End Select
            If adoC!pag_Motivo = 7 Then 'los recargo
                    If adoC!pag_mon = "P" Or adoC!pag_mon = "" Then
                        adoD!debe = Format(adoC!pag_Valor, "#,#0.00")
                        adoD!fdebe = adoC!pag_Valor
                        adoD!fdife = adoC!pag_Valor
                    Else
                        adoD!debe = Format(adoC!pag_ValME, "#,#0.00")
                        adoD!fdebe = adoC!pag_ValME
                       adoD!fdife = adoC!pag_ValME
                      End If
            Else
                    If adoC!pag_mon = "P" Or adoC!pag_mon = "" Then
                        adoD!haber = Format(adoC!pag_Valor, "#,#0.00")
                        adoD!fhaber = adoC!pag_Valor
                        adoD!fdife = -adoC!pag_Valor
                      Else
                        adoD!haber = Format(adoC!pag_ValME, "#,#0.00")
                        adoD!fhaber = adoC!pag_ValME
                         adoD!fdife = -adoC!pag_ValME
                      End If
            End If
            adoC.MoveNext
        Loop
        adoD.Update
        'Set fjMome.DataGrid1.DataSource = adoD     'cOrd.adoOrdenes
        'fjMome.Show
        'Exit Sub
    End If
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Movimientos"
        GoTo miFinal
    End If

    adoD.Sort = "vencim"
    'Set DataGrid1.DataSource = adoC
  drInfoHistSocio.Caption = "Informe de  " & sNombre & "   Nro:" & Text1.Text

  drInfoHistSocio.Title = "Informe de:  " & sNombre & vbCrLf & " Socio Nro: " & Text1.Text & "   al " & Date

  Set drInfoHistSocio.DataSource = adoD
  drInfoHistSocio.DataMember = ""

  drInfoHistSocio.Sections(3).Controls(1).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(1).DataField = "Vencim"
  drInfoHistSocio.Sections(3).Controls(2).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(2).DataField = "Tipo"
  drInfoHistSocio.Sections(3).Controls(3).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(3).DataField = "Moned"
  drInfoHistSocio.Sections(3).Controls(4).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(4).DataField = "NoOrden"
  drInfoHistSocio.Sections(3).Controls(5).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(5).DataField = "debe"
  drInfoHistSocio.Sections(3).Controls(6).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(6).DataField = "haber"
    drInfoHistSocio.Sections(3).Controls(7).DataMember = ""
  drInfoHistSocio.Sections(3).Controls(7).DataField = "nComercio"
'totales
  drInfoHistSocio.Sections(5).Controls(1).DataMember = ""
  drInfoHistSocio.Sections(5).Controls(1).DataField = "fdebe"
   drInfoHistSocio.Sections(5).Controls(2).DataMember = ""
  drInfoHistSocio.Sections(5).Controls(2).DataField = "fhaber"
   drInfoHistSocio.Sections(5).Controls(3).DataMember = ""
  drInfoHistSocio.Sections(5).Controls(3).DataField = "fdife"
  
  drInfoHistSocio.Refresh
        Screen.MousePointer = vbDefault
  
  drInfoHistSocio.Show
miFinal:
    cOrd.msTermina
    cPag.mfCierraPagos
    
    If adoCom.State = adStateOpen Then adoCom.Close
    Set adoCom = Nothing
    If adoC.State = adStateOpen Then adoC.Close
    Set adoC = Nothing
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
End Sub


Private Sub mInformeCuentaSocios()
    Dim sMome As String
    Dim sM As String
    
    Dim t1 As String
    Dim t2 As String
    Dim t3 As String
    
    Dim s1 As Long
    Dim s2 As Long
    Dim s3 As Long
    Dim s4 As Long
    Dim s5 As Long
    Dim S6 As Long
    Dim S7 As Long
    Dim s8 As Long
    Dim s9 As Long
    Dim s10 As Long
    
    'abre solo para conectar un data source
    If adoClie.State = adStateOpen Then adoC.Close
    adoClie.Open "Select * from tbl_Parametros;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText

    If adoC.State = adStateOpen Then adoC.Close
    sM = mfInvierteMes(Text1.Text)
    sMome = "SELECT * fROM TBL_Socios WHERE Fech_Ing <=#" & sM & "#;"
    adoC.Open sMome, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoC.RecordCount < 1 Then
        MsgBox "Sin registros"
         If adoC.State = adStateOpen Then adoC.Close
        Exit Sub
    End If
    s1 = adoC.RecordCount
    t1 = t1 & vbCrLf & "Total Socios:...................."
    t2 = t2 & vbCrLf & Format(s1, "#######")
    
    
    
    'total activos
    adoC.Filter = "codCatSoc = 1"
    s2 = adoC.RecordCount
    t1 = t1 & vbCrLf & vbCrLf & "Cat.Activos:....................."
    t2 = t2 & vbCrLf & vbCrLf & Format(s2, "#######")
    
    'total Honorarios
    adoC.Filter = "codCatSoc = 2"
    s3 = adoC.RecordCount
    t1 = t1 & vbCrLf & "Honorarios:......................"
    t2 = t2 & vbCrLf & Format(s3, "#######")
 
    'total Cooperadores
    adoC.Filter = "codCatSoc = 3"
    s4 = adoC.RecordCount
    t1 = t1 & vbCrLf & "Cooperadores:...................."
    t2 = t2 & vbCrLf & Format(s4, "#######")
    
    'otras categorias
    s5 = s1 - s2 - s3 - s4
    t1 = t1 & vbCrLf & "Otras Categorías:................"
    t2 = t2 & vbCrLf & Format(s5, "#######")
   
    'total Sit Lab activos
    adoC.Filter = "codSitLab = 1"
    S6 = adoC.RecordCount
    t1 = t1 & vbCrLf & vbCrLf & "Sit.Lab.Activos:........"
    t2 = t2 & vbCrLf & vbCrLf & Format(S6, "#######")
    
    'total Sit Lab Comisión
    adoC.Filter = "codSitLab = 2"
    S7 = adoC.RecordCount
    t1 = t1 & vbCrLf & "Sit.Lab.Comisión:................"
    t2 = t2 & vbCrLf & Format(S7, "#######")
    
    'total Sit Lab Retirados
    adoC.Filter = "codSitLab = 3"
    s8 = adoC.RecordCount
    t1 = t1 & vbCrLf & "Sit.Lab.Retirados:..............."
    t2 = t2 & vbCrLf & Format(s8, "#######")
  
    'total Sit Lab Pensionistas
    adoC.Filter = "codSitLab = 4"
    s9 = adoC.RecordCount
    t1 = t1 & vbCrLf & "Sit.Lab.Pensionistas:............"
    t2 = t2 & vbCrLf & Format(s9, "#######")
    
    'total Sit Sin clasificar
    s10 = s1 - S6 - S7 - s8 - s9
    t1 = t1 & vbCrLf & "Sit.Lab.Otras Cat:..............."
    t2 = t2 & vbCrLf & Format(s10, "#######")

    drCierreMesResumen.Caption = "Cantidad de Socios al " & Text1.Text
    drCierreMesResumen.Title = "Cantidad de Socios al " & Text1.Text
    drCierreMesResumen.Sections(1).Controls(2).Caption = t1
    drCierreMesResumen.Sections(1).Controls(4).Caption = t2
    drCierreMesResumen.Sections(1).Controls(5).Caption = t3
    drCierreMesResumen.Sections(1).Controls(7).Caption = "" 'anula el que dice pesos
    drCierreMesResumen.Sections(1).Controls(8).Caption = "" 'anula el que dice moneda extranjera
    drCierreMesResumen.Sections(1).Controls(9).Caption = Date
    
    Set drCierreMesResumen.DataSource = adoClie
    drCierreMesResumen.Show
    
End Sub


'===============================================================
'===============================================================
Private Sub mInfoResumen()
Dim sM As String
Dim nM As Integer
Dim sM1 As String           'fecha inicial
Dim sM2 As String           'fecha final




Dim sFechaFinPresup As Date
Dim sFechaIniPresup As Date

PBar.Visible = True
Screen.MousePointer = vbHourglass

Call CeraTotales

'0) Coloca las fechas de inicio y fin
nM = Left(Text1.Text, 2)
PBar.Max = 20
PBar.Min = 0
If nM = 0 Then
    sFechaFinPresup = Format(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac, "short date")
ElseIf nM > 12 Or nM < 1 Then
    Text1.SetFocus
    Exit Sub
Else
    sFechaFinPresup = Format(vpnPrspHst & "/" & Left(Text1.Text, 2) & "/" & Right(Text1.Text, 4), "short date")
End If
If CInt(Left(Text1.Text, 2)) = 1 Then
    sFechaIniPresup = Format(Day(sFechaFinPresup) + 1 & "/12/" & Year(sFechaFinPresup) - 1, "short date")

Else
    Debug.Print Day(sFechaFinPresup) + 1
    Debug.Print Month(sFechaFinPresup) - 1
    Debug.Print Year(sFechaFinPresup)
    sFechaIniPresup = Format(Day(sFechaFinPresup) + 1 & "/" & Month(sFechaFinPresup) - 1 & "/" & Year(sFechaFinPresup), "short date")
End If

'1) abre el resumen
'Open "tmpResumen.tmp" For Output As 2
'Print #2, "Resumen de Movimientos desde el: " & sFechaIniPresup & _
'    " hasta el: " & sFechaFinPresup


'2) Cuenta las ordenes emitidas en esa fecha
sM1 = mfInvierteMes(CStr(sFechaIniPresup))
sM2 = mfInvierteMes(CStr(sFechaFinPresup))

sM = "select * FROM tbl_Ordenes" & _
            " WHERE ORD_FEmis BETWEEN #" & _
            sM1 & "# AND #" & sM2 & "#;"
If adoD.State = adStateOpen Then adoD.Close
adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
'MsgBox adoD.RecordCount
PBar.Value = 2

tT1 = ""
tT2 = ""
tT3 = ""
tT4 = ""

sTot1 = 0
sTot2 = 0
sTot3 = 0
Label5.Caption = "3. Sumando..."
Label5.Refresh
'3) Ordenes carnicería
mColoca1 "ORD_NroCom=115", "Carnicería..........."
PBar.Value = PBar.Value + 2


'4) Ordenes vales
mColoca1 "ORD_NroCom=190", "Vales................"
PBar.Value = PBar.Value + 2

'5) Cuota
mColoca1 "ORD_NroOrden=1", "Cuotas..............."
PBar.Value = PBar.Value + 2

'6) Ayuda
mColoca1 "ORD_NroOrden=2", "Ayuda................"
PBar.Value = PBar.Value + 2

'7) Sub totales
tT1 = tT1 & "Sub-Total............" & vbCrLf
tT2 = tT2 & Format(sTot1, "#,#0.00") & vbCrLf
tT3 = tT3 & Format(sTot2, "0.00") & vbCrLf
tT4 = tT4 & MiFormato2(sTot3, 6) & vbCrLf

'8) Ordenes comunes
adoD.Filter = ""
nA1 = adoD.RecordCount
If nA1 > 0 Then
    Call mSuma(1)
Else
    sMome1 = 0
    sMome2 = 0
End If

tT1 = tT1 & "Ordenes.............." & vbCrLf
tT2 = tT2 & Format(sMome1 - sTot1, "#,#0.00") & vbCrLf
tT3 = tT3 & Format(sMome2 - sTot2, "#,#0.00") & vbCrLf
tT4 = tT4 & MiFormato2(nA1 - sTot3, 6) & vbCrLf
PBar.Value = PBar.Value + 2

'9) Totales
tT1 = tT1 & "Totales.............." & vbCrLf & vbCrLf
tT2 = tT2 & Format(sMome1, "#,#0.00") & vbCrLf & vbCrLf
tT3 = tT3 & Format(sMome2, "0.00") & vbCrLf & vbCrLf
tT4 = tT4 & MiFormato2(nA1, 6) & vbCrLf & vbCrLf

'10) Cobros
sM = "select * FROM tbl_Pagos" & _
            " WHERE Pag_Fecha BETWEEN #" & _
            sM1 & "# AND #" & sM2 & "#;"
If adoC.State = adStateOpen Then adoC.Close
adoC.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText '

mColocaPago "pag_Motivo=5 OR pag_Motivo=6", "Cobros..............."
PBar.Value = PBar.Value + 2
mColocaPago "pag_Motivo=4", "Ordenes Anuladas....."
PBar.Value = PBar.Value + 2
mColocaPago "pag_Motivo=8", "Cobros Anulados......"
PBar.Value = PBar.Value + 2
mColocaPago "pag_Motivo=7", "Recargo Generado....."
PBar.Value = PBar.Value + 2




'11)Total de deuda hasta esa fecha
sM = "select * FROM tbl_Ordenes" & _
            " WHERE ORD_FEmis <= #" & sM2 & "# AND" & _
            " ord_Cerro < #01/01/1901#;"
If adoE.State = adStateOpen Then adoE.Close
adoE.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
PBar.Min = 0
PBar.Max = adoE.RecordCount

adoE.MoveFirst
sMome1 = 0
sMome2 = 0
Label5.Caption = "2. Totales..."
Label5.Refresh

Do While Not adoE.EOF
      If adoE.AbsolutePosition / 200 Then
        PBar.Value = adoE.AbsolutePosition
    End If
    If adoE!ord_Mon = "P" Then
        sMome1 = sMome1 + (adoE!ORD_PLAN - adoE!ord_ctasPagas) * adoE!ord_cuota + adoE!ord_Recarg - adoE!ord_EntCta
    Else
        sMome2 = sMome2 + (adoE!ORD_PLAN - adoE!ord_ctasPagas) * adoE!ord_mecuota + adoE!ord_Recarg - adoE!ord_EntCta
    End If
    adoE.MoveNext
Loop
tT1 = tT1 & "Deuda ...." & vbCrLf
tT2 = tT2 & Format(sMome1, "#,#0.00") & vbCrLf
tT3 = tT3 & Format(sMome2, "#,#0.00") & vbCrLf
tT4 = tT4 & "" & vbCrLf

DoEvents

'---------------------------------------------------------------
'10) dETALLE: Cuenta de otra manera
'OJO: LAS ORDENES CUYO SOCIO NO ESTÁ EN TBL_SOCIOS, NO APARECEN
sM = "select * FROM tbl_Ordenes INNER JOIN tbl_Socios" & _
            " ON tbl_Socios.NroSoc = tbl_Ordenes.ORD_NroSoc" & _
            " WHERE ORD_FEmis BETWEEN #" & _
            sM1 & "# AND #" & sM2 & "#;"
If adoD.State = adStateOpen Then adoD.Close
adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
Dim mInd1 As Integer        'indice del grupo: 0=carniceria 1=vales 2=cuota 3=ayuda 4=ordenes
Dim mInd2 As Integer        '0=ninguno 1=activo 2=comision3=retirado 4=pensionista
Dim mInd3 As Integer        '0=pesos 1=dolar 2=real 3=austral
Dim sValor As Single
PBar.Min = 0
PBar.Max = adoD.RecordCount
Label5.Caption = "1. Ordenando..."
Label5.Refresh
adoD.MoveFirst
Do While Not adoD.EOF
    If adoD.AbsolutePosition / 200 Then
        PBar.Value = adoD.AbsolutePosition
    End If
    Select Case adoD!ORD_NroCom
        Case 115
            mInd1 = 0
        Case 190
            mInd1 = 1
        Case 0
            If adoD!ord_NroOrden = 1 Then         'es una cuota
                mInd1 = 2
            Else                            'es una ayuda
                mInd1 = 3
            End If
        Case Else
            mInd1 = 4
    End Select
    Select Case adoD!CodSitLab
        Case 0
            mInd2 = 0
        Case 1
            mInd2 = 1
        Case 2
            mInd2 = 2
        Case 3
            mInd2 = 3
        Case 4
            mInd2 = 4
        Case Else
            MsgBox "ERROR 6666: El socio " & adoD!ord_nrosoc & " Orden; " & adoD!ord_NroOrden & " Tiene Situac. Laboral:" & adoD!ORD_CodSitLab
    End Select
    Select Case adoD!ord_Mon
        Case "P"
            mInd3 = 0
            sValor = adoD!ord_cuota * adoD!ORD_PLAN
        Case "D"
            mInd3 = 1
            sValor = adoD!ord_mecuota * adoD!ORD_PLAN
        Case "R"
            mInd3 = 2
            sValor = adoD!ord_mecuota * adoD!ORD_PLAN
        Case "A"
            mInd3 = 3
            sValor = adoD!ord_mecuota * adoD!ORD_PLAN
        Case Else
            MsgBox "ERROR 6666: El socio " & adoD!ord_nrosoc & " Orden; " & adoD!ord_NroOrden & " Tiene Moneda:" & adoD!ord_Mon
    End Select
    sT1(mInd1, mInd2, mInd3) = sT1(mInd1, mInd2, mInd3) + sValor
    sT2(mInd1, mInd2) = sT2(mInd1, mInd2) + 1
   adoD.MoveNext
Loop

tT1 = tT1 & vbCrLf & "Detalles" & vbCrLf
tT2 = tT2 & vbCrLf & vbCrLf
tT3 = tT3 & vbCrLf & vbCrLf
tT4 = tT4 & vbCrLf & vbCrLf

mColoca2 0, 0, "Carnicer.Otros........"
mColoca2 0, 1, "Carnicer.Activos......"
mColoca2 0, 2, "Carnicer.Comisión....."
mColoca2 0, 3, "Carnicer.Retirados...."
mColoca2 0, 4, "Carnicer.Pensionistas."
mTotal 0, "Carnicer.Total........"
mColoca2 1, 0, "Vales....Otros........"
mColoca2 1, 1, "Vales....Activos......"
mColoca2 1, 2, "Vales....Comisión....."
mColoca2 1, 3, "Vales....Retirados...."
mColoca2 1, 4, "Vales....Pensionistas."
mTotal 1, "Vales.Total..........."
mColoca2 2, 0, "Cuotas...Otros........"
mColoca2 2, 1, "Cuotas...Activos......"
mColoca2 2, 2, "Cuotas...Comisión....."
mColoca2 2, 3, "Cuotas...Retirados...."
mColoca2 2, 4, "Cuotas...Pensionistas."
mTotal 2, "Cuota.Total..........."
mColoca2 3, 0, "Ayuda....Otros........"
mColoca2 3, 1, "Ayuda....Activos......"
mColoca2 3, 2, "Ayuda....Comisión....."
mColoca2 3, 3, "Ayuda....Retirados...."
mColoca2 3, 4, "Ayuda....Pensionistas."
mTotal 3, "Ayuda.Total..........."
mColoca2 4, 0, "Ordenes..Otros........"
mColoca2 4, 1, "Ordenes..Activos......"
mColoca2 4, 2, "Ordenes..Comisión....."
mColoca2 4, 3, "Ordenes..Retirados...."
mColoca2 4, 4, "Ordenes..Pensionistas."
mTotal 4, "Ordenes.Total........."




'0) Final
Label5.Caption = ""
Label5.Refresh
If adoD.State = adStateOpen Then adoD.Close
Set adoD = Nothing
'Close 2
'vptReporte = "tmpResumen.tmp"
Screen.MousePointer = vbDefault

    drResumenMes.Caption = "Resumen Mensual del: " & sFechaIniPresup & " al:" & sFechaFinPresup
    drResumenMes.Title = "Resumen Mensual " & vbCrLf & "del: " & sFechaIniPresup & " al: " & sFechaFinPresup
    drResumenMes.Sections(1).Controls(2).Caption = tT1
    drResumenMes.Sections(1).Controls(4).Caption = tT2
    drResumenMes.Sections(1).Controls(5).Caption = tT3
    drResumenMes.Sections(1).Controls(10).Caption = tT4
    'drResumenMes.Sections(1).Controls(7).Caption = "" 'anula el que dice pesos
    'drResumenMes.Sections(1).Controls(8).Caption = "" 'anula el que dice moneda extranjera
    drResumenMes.Sections(1).Controls(9).Caption = Date
    
    'abre solo para conectar un data source
    If adoClie.State = adStateOpen Then adoClie.Close
    adoClie.Open "Select * from tbl_Parametros;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Set drResumenMes.DataSource = adoClie
    
    drResumenMes.Show
    
End Sub

Private Sub mColoca2(a1 As Integer, a2 As Integer, s1 As String)
If Not sT1(a1, a2, 0) + sT1(a1, a2, 1) + sT1(a1, a2, 2) + sT1(a1, a2, 3) = 0 Then
    tT1 = tT1 & s1 & vbCrLf
    tT2 = tT2 & Format(sT1(a1, a2, 0), "#,#0.00") & vbCrLf
    tT3 = tT3 & Format(sT1(a1, a2, 1) + sT1(a1, a2, 2) + sT1(a1, a2, 3), "#,#0.00") & vbCrLf
    tT4 = tT4 & sT2(a1, a2) & vbCrLf
End If
End Sub


Private Sub mTotal(a1 As Integer, s1 As String)
Dim sSuma1 As Single
Dim sSuma2 As Single
Dim sSuma3 As Long
Dim nM As Byte

For nM = 0 To 4
    sSuma1 = sSuma1 + sT1(a1, nM, 0)
    sSuma2 = sSuma2 + sT1(a1, nM, 1) + sT1(a1, nM, 2) + sT1(a1, nM, 3)
    sSuma3 = sSuma3 + sT2(a1, nM)
Next


If Not sSuma1 + sSuma2 = 0 Then
    tT1 = tT1 & s1 & vbCrLf & vbCrLf
    tT2 = tT2 & Format(sSuma1, "#,#0.00") & vbCrLf & vbCrLf
    tT3 = tT3 & Format(sSuma2, "#,#0.00") & vbCrLf & vbCrLf
    tT4 = tT4 & sSuma3 & vbCrLf & vbCrLf
End If
End Sub

Private Sub CeraTotales()
Dim na As Byte
Dim nb As Byte
Dim nc As Byte
For na = 0 To 4
    For nb = 0 To 4
        For nc = 0 To 3
            sT1(na, nb, nc) = 0
        Next nc
    Next nb
Next na
For na = 0 To 4
    For nb = 0 To 4
            sT2(na, nb) = 0
    Next nb
Next na

End Sub

Private Sub mColoca1(sPrm1 As String, sPrm2 As String)
adoD.Filter = sPrm1
nA1 = adoD.RecordCount
If nA1 > 0 Then
    Call mSuma(1)
Else
    sMome1 = 0
    sMome2 = 0
End If

sTot1 = sTot1 + sMome1
sTot2 = sTot2 + sMome2
sTot3 = sTot3 + nA1
tT1 = tT1 & sPrm2 & vbCrLf
tT2 = tT2 & Format(sMome1, "#,#0.00") & vbCrLf
tT3 = tT3 & Format(sMome2, "#,#0.00") & vbCrLf
tT4 = tT4 & MiFormato2(nA1, 6) & vbCrLf
End Sub



Private Sub mColocaPago(sPrm1 As String, sPrm2 As String)
adoC.Filter = sPrm1
nA1 = adoC.RecordCount
If nA1 > 0 Then
    Call mSuma(2)
Else
    sMome1 = 0
    sMome2 = 0
End If

sTot1 = sTot1 + sMome1
sTot2 = sTot2 + sMome2
sTot3 = sTot3 + nA1
tT1 = tT1 & sPrm2 & vbCrLf
tT2 = tT2 & Format(sMome1, "#,#0.00") & vbCrLf
tT3 = tT3 & Format(sMome2, "#,#0.00") & vbCrLf
tT4 = tT4 & MiFormato2(nA1, 6) & vbCrLf
End Sub


Private Sub mSuma(nPrm As Byte)
Dim sM5 As Single
Select Case nPrm
Case 1
    sMome1 = 0
    sMome2 = 0
    adoD.MoveFirst
    Do While Not adoD.EOF
        If adoD!ord_Mon = "P" Then
            sMome1 = sMome1 + adoD!ord_cuota * adoD!ORD_PLAN
        Else
            sMome2 = sMome2 + adoD!ord_mecuota * adoD!ORD_PLAN
        End If
        adoD.MoveNext
    Loop
Case 2
    sMome1 = 0
    sMome2 = 0
    adoC.MoveFirst
    Do While Not adoC.EOF
        If adoC!pag_mon = "P" Then
            sMome1 = sMome1 + adoC!pag_Valor
        Else
            sMome2 = sMome2 + adoC!pag_ValME
        End If
        adoC.MoveNext
    Loop
End Select
End Sub



Private Function MiFormato(sPrm As Single) As String
Dim s1 As String
s1 = Format(sPrm, "#,#0.00")
MiFormato = Space(30 - Len(s1)) & s1
End Function




Private Function MiFormato2(lPrm As Long, nPrm As Byte) As String
Dim s1 As String
s1 = CStr(lPrm)
MiFormato2 = Space(nPrm - Len(s1)) & s1
End Function



'=======================================================
Private Sub mInformeHistoriaUnSocio2()
'=======================================================
    Dim sNombre As String
    Screen.MousePointer = vbHourglass
    
    If adoClie.State = adStateOpen Then adoClie.Close
    adoClie.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    
    If adoCom.State = adStateOpen Then adoCom.Close
    adoCom.Open "select * FROM tbl_Comercios ORDER BY codigo;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    
    '6.1) Coloca el nombre
    adoClie.MoveFirst
    adoClie.Find "NroSoc =" & Text1.Text
    If Not adoClie.EOF Then
        sNombre = Left(Trim(adoClie!Apellido) & "  " & _
        Trim(adoClie!nombre), 30) & "  Nro.Cob. " & adoClie!NroCob
    Else
        sNombre = "Desconocido"
    End If
    adoClie.Close
    Set adoClie = Nothing
 
 
    '1)BUSCA LAS ORDENES QUE TIENE EL SOCIO
    cOrd.msInicia
    cOrd.vlNroSoc = Text1.Text
    
    If Not cOrd.fBusca2TodasLasOrdenesUnSocio Then      'HUBO PROBLEMAS
        MsgBox "2328s: Problemas al Buscar Ordenes (1) "
        GoTo miFinal
    End If
    
    'If cOrd.adoOrdenes.RecordCount = 0 Then
    '    MsgBox "No tiene Ordenes"
    '    GoTo miPagos
    'End If
    
    'Set fjMome.DataGrid1.DataSource = cOrd.adoOrdenes
    'fjMome.Show vbModal
    'MsgBox "hola"
  
    '2.Prepara el ado
    
    cOrd.ms2PreparaOrdenesAPagarEnAdoM2
    If cOrd.adoM2.RecordCount < 1 Then
        Exit Sub
    End If
    Set adoC = cOrd.adoM2
    
    'Set fjMome.DataGrid1.DataSource = adoC     'cOrd.adoOrdenes
    'fjMome.Show vbmodal
    'MsgBox "hola"
    'Exit Sub
    
    '3) crea la tabla virtual adoD
    Set adoD.ActiveConnection = adoConn
    Set adoD = New ADODB.Recordset
    Dim nM As Integer
    Dim mNomb As String
    Dim mTipo
    Dim mTam As Long
    '3.1) con los mismos campos
    For nM = 0 To adoC.Fields.Count - 1
        mNomb = adoC.Fields(nM).Name
        mTipo = adoC.Fields(nM).Type
        mTam = adoC.Fields(nM).DefinedSize
        adoD.Fields.Append mNomb, mTipo, mTam
    Next
    '3.2) con campos nuevos
    adoD.Fields.Append "nComercio", adChar, 30
    adoD.Fields.Append "Tipo", adChar, 30
    adoD.Fields.Append "Mensaje", adChar, 30
    adoD.Fields.Append "Haber", adChar, 30              'es un char
    adoD.Fields.Append "sFecha", adChar, 10
    adoD.Fields.Append "Debe", adChar, 30    'es un char
    adoD.Fields.Append "STot", adChar, 15
     
    adoD.CursorType = adOpenDynamic
    adoD.LockType = adLockOptimistic
    adoD.Open
    '4) LLena los campos
     
    adoC.MoveFirst
    mTam = adoC.Fields.Count - 1
    Do While Not adoC.EOF
        adoD.AddNew
        'los campos comunes
        For nM = 0 To mTam
            adoD(nM) = adoC(nM)
        Next
        'los campos especiales
        'tipo
        Select Case adoC("noorden")
            Case 0
                adoD("Tipo") = ""
            Case 1
                adoD("Tipo") = "Cuota"
            Case 2
               adoD("Tipo") = "Ayuda"
            Case 3
               adoD("Tipo") = "Recargo"
            Case Else
               adoD("Tipo") = ""
        End Select
        '6.4) Coloca el Comercio
        adoCom.MoveFirst
        adoCom.Find "Codigo =" & adoC!comercio
        If Not adoCom.EOF Then
            adoD!nComercio = "" & Left(Trim(adoCom!NombCom), 30)
        Else
            adoD!nComercio = ""
        End If
        adoD!valorp = adoD!valorp + adoD!valorME
        adoD!haber = ""
        adoD!debe = Format(adoD!valorp, "#,#0.00")
        adoD("Mensaje") = Trim(adoD("tipo") & adoD("ncomercio"))

     adoC.MoveNext
    Loop
    adoD.Update
    adoC.Close
    'Set fjMome.DataGrid1.DataSource = adoD     'cOrd.adoOrdenes
    'fjMome.Show
    'Exit Sub
  

    If adoD.RecordCount < 1 Then
        MsgBox "Sin Movimientos"
        GoTo miFinal
    End If
   'Coloca subtotales por dia de vencimiento
  Dim mSTot As Double
  Dim sfecha As String
  mSTot = 0
   adoD.Sort = "vencim"

  adoD.MoveFirst
  sfecha = adoD!vencim
  adoD!sfecha = CStr(adoD!vencim)
  
  Do While Not adoD.EOF
        'cambio de fecha de vencimiento
        If Not adoD!vencim = sfecha Then
            adoD.MovePrevious
            adoD!sTot = Format(mSTot, "#,#0.00")
            adoD.MoveNext
            mSTot = 0
            sfecha = adoD!vencim
            adoD!sfecha = CStr(adoD!vencim)
        End If
        mSTot = mSTot + adoD!valorp
        adoD.MoveNext
    Loop
adoD.MovePrevious
adoD!sTot = Format(mSTot, "#,#0.00")

     'Set DataGrid1.DataSource = adoC
  drInfoSocio2.Caption = "Informe de  " & sNombre & "   Nro:" & Text1.Text

  drInfoSocio2.Title = "Informe de:  " & sNombre & vbCrLf & " Socio Nro: " & Text1.Text & "   al " & Date

  Set drInfoSocio2.DataSource = adoD
  drInfoSocio2.DataMember = ""

  drInfoSocio2.Sections(3).Controls(2).DataMember = ""
  drInfoSocio2.Sections(3).Controls(2).DataField = "sfecha"
  drInfoSocio2.Sections(3).Controls(1).DataMember = ""
  drInfoSocio2.Sections(3).Controls(1).DataField = "noorden"
  drInfoSocio2.Sections(3).Controls(3).DataMember = ""
  drInfoSocio2.Sections(3).Controls(3).DataField = "nodepend"
  drInfoSocio2.Sections(3).Controls(4).DataMember = ""
  drInfoSocio2.Sections(3).Controls(4).DataField = "mensaje"
  drInfoSocio2.Sections(3).Controls(5).DataMember = ""
  drInfoSocio2.Sections(3).Controls(5).DataField = "valorp"
  drInfoSocio2.Sections(3).Controls(6).DataMember = ""
  drInfoSocio2.Sections(3).Controls(6).DataField = "sinpagar"
  drInfoSocio2.Sections(3).Controls(7).DataMember = ""
  drInfoSocio2.Sections(3).Controls(7).DataField = "moned"
  drInfoSocio2.Sections(3).Controls(8).DataMember = ""
  drInfoSocio2.Sections(3).Controls(8).DataField = "stot"
  drInfoSocio2.Sections(3).Controls(9).DataMember = ""
  drInfoSocio2.Sections(3).Controls(9).DataField = "cuota"
  'totales
  drInfoSocio2.Sections(5).Controls(1).DataMember = ""
  drInfoSocio2.Sections(5).Controls(1).DataField = "valorp"
  drInfoSocio2.Sections(5).Controls(3).DataMember = ""
  drInfoSocio2.Sections(5).Controls(3).DataField = "sinPagar"
  
  drInfoSocio2.Refresh
  Screen.MousePointer = vbDefault
   drInfoSocio2.Show
miFinal:
    cOrd.msTermina
    cPag.mfCierraPagos
    
    If adoCom.State = adStateOpen Then adoCom.Close
    Set adoCom = Nothing
    If adoC.State = adStateOpen Then adoC.Close
    Set adoC = Nothing
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
End Sub




'=======================================================
Private Sub mInfoCobros2()
'=======================================================
    'total de pagos por socio.
    Dim sM1 As String
    Dim sM2 As String
    
      
    Screen.MousePointer = vbHourglass
    PBar.Visible = False
    Label5.Caption = "Espere..."
    Label5.Refresh
    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String
    
     sM = "select *,space(25) as apellido, space(25) as nombre FROM tbl_Pagos" & _
           " WHERE not pag_Motivo = 4 AND" & _
           " NOT pag_Motivo = 7 AND" & _
            " pag_Fecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
    adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
     
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
  
  
      'coloca los nombres
    Label5.Caption = "Nombres..."
    Label5.Refresh
    cSocio.mfAbreTablaSociosOrdenSocio
    adoD.MoveFirst
    Do While Not adoD.EOF
        cSocio.vlNroSoc = adoD!pag_NroSoc
        If Not cSocio.mfBuscaSocio Then
            adoD!Apellido = "NO ENCONTRADO"
            adoD!nombre = ""
        Else
            adoD!Apellido = cSocio.vsApellido
            adoD!nombre = cSocio.vsNombre
            adoD.Update
        End If
        adoD.MoveNext
    Loop

    adoD.Sort = "pag_NroPago, pag_NroSoc"
    'adoD.Filter = "pag_motivo <> 7"     'sin recargo
    
    'Crea un ado virtual
    Label5.Caption = "Creando ado...."
    Label5.Refresh
    Set adoE.ActiveConnection = adoConn
    Set adoE = New ADODB.Recordset
    adoE.Fields.Append "Fecha", adDate
    adoE.Fields.Append "NRecibo", adInteger, 2
    adoE.Fields.Append "NSocio", adInteger, 2
    adoE.Fields.Append "Nombre", adChar, 50
    adoE.Fields.Append "ValorP", adSingle
    adoE.Fields.Append "ValorME", adSingle
     
    adoE.CursorType = adOpenDynamic
    adoE.LockType = adLockOptimistic
    adoE.Open



    
    'Prepara el informe
    Label5.Caption = "Prepara...."
    Label5.Refresh
    Dim nFech As Date
    Dim nRec As Long
    Dim nSoc As Long
    Dim sNomb As String
    Dim sTotP As Single
    Dim sTotD As Single
    
    PBar.Visible = True
    PBar.Min = 0
    PBar.Max = adoD.RecordCount
    
    adoD.MoveFirst
    nFech = adoD!pag_Fecha
    nRec = adoD!pag_NroPago
    nSoc = adoD!pag_NroSoc
    sNomb = adoD!Apellido & " " & adoD!nombre
    sTotP = 0
    sTotD = 0
    
    Do While Not adoD.EOF
        If adoD.AbsolutePosition Mod 100 = 0 Then PBar.Value = adoD.AbsolutePosition
        If Not adoD!pag_NroPago = nRec Or Not adoD!pag_NroSoc = nSoc Then
            Call AgregaRegistro(nFech, nRec, nSoc, sNomb, sTotP, sTotD)
            nFech = adoD!pag_Fecha
            nRec = adoD!pag_NroPago
            nSoc = adoD!pag_NroSoc
            sNomb = adoD!Apellido & " " & adoD!nombre
            sTotP = 0
            sTotD = 0
        End If
        If adoD!pag_mon = "P" Then
            sTotP = sTotP + adoD!pag_Valor
        Else
            sTotD = sTotD + adoD!pag_ValME
        End If
        adoD.MoveNext
    Loop
    Call AgregaRegistro(nFech, nRec, nSoc, sNomb, sTotP, sTotD)
  
     
    ' Muestra el informe de datos
    'un solo dia
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfCobros2.Caption = "Resumen de Cobros del día: " & Text1.Text
        drInfCobros2.Title = "Resumen de Cobros " & vbCrLf & "del día: " & Text1.Text
     'varios dias
    Else
        drInfCobros2.Caption = "Resumen de Cobros desde: " & Text1.Text & "  hasta: " & Text2.Text
        drInfCobros2.Title = "Resumen de Cobros" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
    End If
    
    Set drInfCobros2.DataSource = adoE
    drInfCobros2.DataMember = ""
    
    drInfCobros2.Sections(3).Controls(1).DataMember = ""
    drInfCobros2.Sections(3).Controls(1).DataField = "Fecha"
    drInfCobros2.Sections(3).Controls(2).DataMember = ""
    drInfCobros2.Sections(3).Controls(2).DataField = "NRecibo"
    drInfCobros2.Sections(3).Controls(3).DataMember = ""
    drInfCobros2.Sections(3).Controls(3).DataField = "NSocio"
    drInfCobros2.Sections(3).Controls(4).DataMember = ""
    drInfCobros2.Sections(3).Controls(4).DataField = "Nombre"
    drInfCobros2.Sections(3).Controls(5).DataMember = ""
    drInfCobros2.Sections(3).Controls(5).DataField = "ValorP"
    drInfCobros2.Sections(3).Controls(6).DataMember = ""
    drInfCobros2.Sections(3).Controls(6).DataField = "ValorMe"
    'totales
    drInfCobros2.Sections(5).Controls(1).DataMember = ""
    drInfCobros2.Sections(5).Controls(1).DataField = "ValorP"
    drInfCobros2.Sections(5).Controls(2).DataMember = ""
    drInfCobros2.Sections(5).Controls(2).DataField = "ValorME"
    
    drInfCobros2.Refresh
    Screen.MousePointer = vbDefault
    Label5.Caption = ""
    Label5.Refresh

    drInfCobros2.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
   If adoE.State = adStateOpen Then adoE.Close
    Set adoE = Nothing
    Unload Me
End Sub
  
Private Sub AgregaRegistro(sF As Date, lR As Long, lS As Long, tN As String, sP As Single, sd As Single)
adoE.AddNew
adoE!Fecha = sF
adoE!NRecibo = lR
adoE!nSocio = lS
adoE!nombre = Left(tN, 50)
adoE!valorp = sP
adoE!valorME = sd
adoE.Update
End Sub



'=======================================
Private Sub mInfoResumen2(nPrm As Byte)
'=======================================
'Solicitudes Discriminados por grupos
'nprm =1 para Jefatura
'nprm = 2 para Retirados
'nprm = 3 Todas las deudas
'           pp_NroCom       pp_NroOrden     ST1(este)
'Vales      190                             1,1,0
'Carnicec   115                             1,1,1
'Cuota                          1           1,1,2
'Ayuda                          2           1,1,3
'Credito                                    1,1,4
'Total                                      1,2,0

Dim nMes As String      'aaaamm
Dim nM As String


CeraTotales
PBar.Visible = True
Screen.MousePointer = vbHourglass

'0) Coloca las fechas de inicio y fin
        If nPrm = 1 Or nPrm = 2 Then
            nM = Left(Text1.Text, 2)
            If nM = 0 Then
                nMes = vpnAñoOperac & IIf(vpbMesOperac > 9, "", "0") & vpbMesOperac
            ElseIf nM > 12 Or nM < 1 Then
                Text1.SetFocus
                Exit Sub
            Else
                nMes = Right(Text1.Text, 4) & Left(Text1.Text, 2)
            End If
        End If
'1) lista la deuda solicitada a Jefatura
        Dim sCadena As String
        If adoC.State = adStateOpen Then adoC.Close
        If nPrm = 1 Then
             sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
                "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc WHERE " & _
                "tbl_socios.codcatsoc = 1 AND tbl_socios.codSitLab = 1 AND " & _
                "tbl_prepago.pp_presup ='" & nMes & "' ORDER BY tbl_prepago.pp_nrosoc;"
        ElseIf nPrm = 2 Then     'centro: retirados y pensionistas
            sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
                "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc WHERE " & _
                "(tbl_socios.codSitLab = 3 OR tbl_socios.codSitLab = 4) AND " & _
                "tbl_prepago.pp_presup ='" & nMes & "' ORDER BY pp_nrosoc;"
        ElseIf nPrm = 3 Then
               sCadena = "SELECT * FROM tbl_ordenes INNER JOIN tbl_socios " & _
                "ON  tbl_socios.nrosoc = tbl_ordenes.ord_NroSoc " & _
                "WHERE year(ord_cerro) < 1901 " & _
                "ORDER BY ord_NroSoc;"
    
        End If
        adoC.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
        If adoC.RecordCount < 1 Then
            MsgBox "Sin Socios"
            Exit Sub
        End If
        
'2) Crea una tabla nueva
        Set adoD.ActiveConnection = adoConn
        Set adoD = New ADODB.Recordset
        adoD.Fields.Append "Socio", adInteger, 2
        adoD.Fields.Append "Nombre", adChar, 30
        adoD.Fields.Append "Credito", adSingle
        adoD.Fields.Append "Carnic", adSingle
        adoD.Fields.Append "Vales", adSingle
        adoD.Fields.Append "Cuota", adSingle
        adoD.Fields.Append "Ayuda", adSingle
        adoD.Fields.Append "Total", adSingle
        'adoD.Fields.Append "NCobro", adInteger, 2
         
        adoD.CursorType = adOpenDynamic
        adoD.LockType = adLockOptimistic
        adoD.Open
        
'2) Organiza la deuda
        Dim lSoc As Long
        Dim sNmb As String

        Screen.MousePointer = vbHourglass
        PBar.Visible = True
        PBar.Min = 0
        PBar.Max = adoC.RecordCount
        
        Select Case nPrm
            Case 1, 2
                adoC.MoveFirst
                lSoc = adoC!pp_NroSoc
                sNmb = adoC!Apellido & " " & adoC!nombre
                Do While Not adoC.EOF
                        If Not adoC!pp_NroSoc = lSoc Then
                            Rsmn2GrabaDatos lSoc, sNmb
                            lSoc = adoC!pp_NroSoc
                            sNmb = adoC!Apellido & " " & adoC!nombre
                        End If
                        Rsmn2TomaDatos
                        PBar.Value = adoC.AbsolutePosition
                        adoC.MoveNext
                Loop
                Rsmn2GrabaDatos lSoc, sNmb
            Case 3
                adoC.MoveFirst
                lSoc = adoC!ord_nrosoc
                sNmb = adoC!Apellido & " " & adoC!nombre
                Do While Not adoC.EOF
                        If Not adoC!ord_nrosoc = lSoc Then
                            Rsmn2GrabaDatos lSoc, sNmb
                            lSoc = adoC!ord_nrosoc
                            sNmb = adoC!Apellido & " " & adoC!nombre
                        End If
                        Rsmn2TomaDatos2
                        PBar.Value = adoC.AbsolutePosition
                        adoC.MoveNext
                Loop
                Rsmn2GrabaDatos lSoc, sNmb

        End Select
    
        Screen.MousePointer = vbDefault
        PBar.Visible = False
'3) Muestra Resumen
        If nPrm = 1 Then
            drResumen2.Caption = "Informe de Solicitudes a Jefatura. Presup: " & nMes
            drResumen2.Title = "Informe de Solicitudes a Jefatura. Presup: " & nMes & " Impreso:" & Date
        ElseIf nPrm = 2 Then
            drResumen2.Caption = "Informe de Solicitudes a Centro Policial. Presup: " & nMes
            drResumen2.Title = "Informe de Solicitudes a Centro Policial. Presup: " & nMes & " Impreso:" & Date
        Else
            drResumen2.Caption = "Informe de la deuda de todos los Socios "
            drResumen2.Title = "Informe de la deuda de todos los Socios. Impreso:" & Date
       
        End If
        Set drResumen2.DataSource = adoD
        drResumen2.DataMember = ""
        
        drResumen2.Sections(3).Controls(1).DataMember = ""
        drResumen2.Sections(3).Controls(1).DataField = "socio"
        drResumen2.Sections(3).Controls(2).DataMember = ""
        drResumen2.Sections(3).Controls(2).DataField = "nombre"
        drResumen2.Sections(3).Controls(3).DataMember = ""
        drResumen2.Sections(3).Controls(3).DataField = "vales"
        drResumen2.Sections(3).Controls(4).DataMember = ""
        drResumen2.Sections(3).Controls(4).DataField = "carnic"
        drResumen2.Sections(3).Controls(5).DataMember = ""
        drResumen2.Sections(3).Controls(5).DataField = "cuota"
        drResumen2.Sections(3).Controls(6).DataMember = ""
        drResumen2.Sections(3).Controls(6).DataField = "ayuda"
        drResumen2.Sections(3).Controls(7).DataMember = ""
        drResumen2.Sections(3).Controls(7).DataField = "credito"
        drResumen2.Sections(3).Controls(8).DataMember = ""
        drResumen2.Sections(3).Controls(8).DataField = "total"
        'totales
        drResumen2.Sections(5).Controls(1).DataMember = ""
        drResumen2.Sections(5).Controls(1).DataField = "vales"
        drResumen2.Sections(5).Controls(3).DataMember = ""
        drResumen2.Sections(5).Controls(3).DataField = "carnic"
        drResumen2.Sections(5).Controls(4).DataMember = ""
        drResumen2.Sections(5).Controls(4).DataField = "cuota"
        drResumen2.Sections(5).Controls(5).DataMember = ""
        drResumen2.Sections(5).Controls(5).DataField = "ayuda"
        drResumen2.Sections(5).Controls(6).DataMember = ""
        drResumen2.Sections(5).Controls(6).DataField = "credito"
        drResumen2.Sections(5).Controls(7).DataMember = ""
        drResumen2.Sections(5).Controls(7).DataField = "total"
        
        drResumen2.Refresh
              Screen.MousePointer = vbDefault
        
        drResumen2.Show
'final
        adoC.Close
        adoD.Close
        Set adoC = Nothing
        Set adoD = Nothing
End Sub
Private Sub Rsmn2TomaDatos()
        If adoC!pp_NroCom = 190 Then
            sT1(1, 1, 0) = sT1(1, 1, 0) + adoC!pp_Valor
       ElseIf adoC!pp_NroCom = 115 Then
            sT1(1, 1, 1) = sT1(1, 1, 1) + adoC!pp_Valor
       ElseIf adoC!pp_NroOrden = 1 Then
            sT1(1, 1, 2) = sT1(1, 1, 2) + adoC!pp_Valor
       ElseIf adoC!pp_NroOrden = 2 Then
            sT1(1, 1, 3) = sT1(1, 1, 3) + adoC!pp_Valor
        Else
            sT1(1, 1, 4) = sT1(1, 1, 4) + adoC!pp_Valor
       End If
       sT1(1, 2, 0) = sT1(1, 2, 0) + adoC!pp_Valor
End Sub
Private Sub Rsmn2TomaDatos2()
        Dim sVal As Single
        
        With adoC
            If Not !ord_Mon = "P" Then
                sVal = (!ord_mecuota * (!ORD_PLAN - !ord_ctasPagas) + !ord_Recarg - !ord_EntCta) * cTC.mfDevuelveCambio(!ord_Mon, Date)
            Else
                sVal = !ord_cuota * (!ORD_PLAN - !ord_ctasPagas) + !ord_Recarg - !ord_EntCta
            End If
        End With
        If adoC!ORD_NroCom = 190 Then
            sT1(1, 1, 0) = sT1(1, 1, 0) + sVal
       ElseIf adoC!ORD_NroCom = 115 Then
            sT1(1, 1, 1) = sT1(1, 1, 1) + sVal
       ElseIf adoC!ord_NroOrden = 1 Then
            sT1(1, 1, 2) = sT1(1, 1, 2) + sVal
       ElseIf adoC!ord_NroOrden = 2 Then
            sT1(1, 1, 3) = sT1(1, 1, 3) + sVal
        Else
            sT1(1, 1, 4) = sT1(1, 1, 4) + sVal
       End If
       sT1(1, 2, 0) = sT1(1, 2, 0) + sVal
End Sub

Private Sub Rsmn2GrabaDatos(plSoc, psNmb)
        adoD.AddNew
        adoD!socio = plSoc
        adoD!nombre = Left(psNmb, 30)
        adoD!vales = sT1(1, 1, 0)
        adoD!carnic = sT1(1, 1, 1)
        adoD!cuota = sT1(1, 1, 2)
        adoD!ayuda = sT1(1, 1, 3)
        adoD!credito = sT1(1, 1, 4)
        adoD!Total = sT1(1, 2, 0)
        sT1(1, 1, 0) = 0
        sT1(1, 1, 1) = 0
        sT1(1, 1, 2) = 0
        sT1(1, 1, 3) = 0
        sT1(1, 1, 4) = 0
        sT1(1, 2, 0) = 0
        adoD.Update
 End Sub






'=======================================================
Private Sub mInfoAdmin1(nPrm As Byte)
'=======================================================
    'nprm=1 pesos
    'nprm = 2 Dólares
    
    Screen.MousePointer = vbHourglass
    '0.Crea una tabla virtual
        Label5.Caption = "7.Creando tabla virtual"
        Label5.Refresh
        Set adoD.ActiveConnection = adoConn
        Set adoD = New ADODB.Recordset
        adoD.Fields.Append "Mes", adChar, 4
        adoD.Fields.Append "Fecha", adChar, 8
        adoD.Fields.Append "Carn1", adSingle
        adoD.Fields.Append "Vale1", adSingle
        adoD.Fields.Append "Orde1", adSingle
        adoD.Fields.Append "Cuot1", adSingle
        adoD.Fields.Append "Ayud1", adSingle
        adoD.Fields.Append "Tota1", adSingle
        adoD.Fields.Append "Carn2", adSingle
        adoD.Fields.Append "Vale2", adSingle
        adoD.Fields.Append "Orde2", adSingle
        adoD.Fields.Append "Cuot2", adSingle
        adoD.Fields.Append "Ayud2", adSingle
        adoD.Fields.Append "Tota2", adSingle
        adoD.Fields.Append "Carn3", adSingle
        adoD.Fields.Append "Vale3", adSingle
        adoD.Fields.Append "Orde3", adSingle
        adoD.Fields.Append "Cuot3", adSingle
        adoD.Fields.Append "Ayud3", adSingle
        adoD.Fields.Append "Tota3", adSingle
         
        adoD.CursorType = adOpenDynamic
        adoD.LockType = adLockOptimistic
        adoD.Sort = "mes"
        adoD.Open
 
    
    '1.Toma todas la ordenes
    cOrd.msInicia
   'todas las ordenes incluso las cerradas y NO la anuladas
    Label5.Caption = "6.Buscando ordenes...."
    Label5.Refresh
    If Not cOrd.fBuscaOrdenesTodosSocios5 Then     'HUBO PROBLEMAS
        MsgBox "2348x: Problemas al Buscar Ordenes"
        GoTo miFinal
    End If
    
    If cOrd.adoOrdenes.RecordCount = 0 Then
        MsgBox "No tiene Ordenes"
        GoTo miFinal
    End If

    '2.Filtra la moneda
    Label5.Caption = "5.Filtrando...."
    Label5.Refresh
    If nPrm = 1 Then
        cOrd.adoOrdenes.Filter = "ord_Mon='P'"
    Else
        cOrd.adoOrdenes.Filter = "ord_Mon='D'"
    End If
    
    '3.Prepara el ado
    Label5.Caption = "4.Preparando por fecha..."
    Label5.Refresh
    cOrd.ms7PreparaOrdenesAPagarEnAdoM2 (nPrm)
    'Set fjMome.DataGrid1.DataSource = cOrd.adoM2
    'fjMome.Show
    'MsgBox "Wait"
    'Exit Sub
    
     '4.Suma Todo
    Call SumaTodo_InfAdm1(1)
    'Set fjMome.DataGrid1.DataSource = adoD
    'fjMome.Show
    
    
    '5.Toma todas la ordenes activas
    cOrd.msTermina
    cOrd.msInicia
   'todas las ordenes NO las cerradas y NO la anuladas
    Label5.Caption = "3.Buscando ordenes...."
    Label5.Refresh
    If Not cOrd.fBuscaOrdenesTodosSocios Then     'HUBO PROBLEMAS
        MsgBox "Error 2348Y: Problemas al Buscar Ordenes"
        GoTo miFinal
    End If
    
    If cOrd.adoOrdenes.RecordCount = 0 Then
        MsgBox "Error 4354: No tiene Ordenes"
        GoTo miFinal
    End If

    '6.Filtra moneda
    Label5.Caption = "2.Filtrando...."
    Label5.Refresh
    If nPrm = 1 Then
        cOrd.adoOrdenes.Filter = "ord_Mon='P'"
        'cOrd.adoOrdenes.Filter = "ord_NroSoc=252 AND ord_Mon='P'"
    Else
        cOrd.adoOrdenes.Filter = "ord_Mon='D'"
        'cOrd.adoOrdenes.Filter = "ord_NroSoc=252 AND ord_Mon='D'"
    End If

    
    '7.Prepara el ado
    Label5.Caption = "1.Preparando por fecha..."
    Label5.Refresh
    cOrd.ms8PreparaOrdenesAPagarEnAdoM2 (nPrm)
    'Set fjMome.DataGrid1.DataSource = cOrd.adoM2
    'fjMome.Show
    'msgbox "Wait"
    'Exit Sub
    
     '8.Suma Todo
    Call SumaTodo_InfAdm1(2)
    'Set fjMome.DataGrid1.DataSource = adoD
    'fjMome.Show
    'MsgBox "ESPERE"
    
    '9. caluclua la diferencia
    With adoD
    .MoveFirst
    Do While Not adoD.EOF
        !carn3 = !carn1 - !carn2
        !vale3 = !vale1 - !vale2
        !orde3 = !orde1 - !orde2
        !cuot3 = !cuot1 - !cuot2
        !ayud3 = !ayud1 - !ayud2
        !tota3 = !tota1 - !tota2
        .Update
        .MoveNext
    Loop
    .MoveFirst
    End With
    If nPrm = 1 Then
        drInfAdmin1.Caption = "Ordenes Emitidas y Ordenes NO Canceladas en Pesos"
        drInfAdmin1.Title = "Ordenes Emitidas y Ordenes NO Canceladas en Pesos" & vbCrLf & "Emitido el: " & Date
    Else
        drInfAdmin1.Caption = "Ordenes Emitidas y Ordenes NO Canceladas en Dólares"
        drInfAdmin1.Title = "Ordenes Emitidas y Ordenes NO Canceladas en Dólares" & vbCrLf & "Emitido el: " & Date
    End If
    Set drInfAdmin1.DataSource = adoD
    drInfAdmin1.DataMember = ""
    
    drInfAdmin1.Sections(3).Controls(1).DataMember = ""
    drInfAdmin1.Sections(3).Controls(1).DataField = "Fecha"
    drInfAdmin1.Sections(3).Controls(2).DataMember = ""
    drInfAdmin1.Sections(3).Controls(2).DataField = "Carn1"
    drInfAdmin1.Sections(3).Controls(3).DataMember = ""
    drInfAdmin1.Sections(3).Controls(3).DataField = "Vale1"
    drInfAdmin1.Sections(3).Controls(4).DataMember = ""
    drInfAdmin1.Sections(3).Controls(4).DataField = "Orde1"
    drInfAdmin1.Sections(3).Controls(5).DataMember = ""
    drInfAdmin1.Sections(3).Controls(5).DataField = "Cuot1"
    drInfAdmin1.Sections(3).Controls(6).DataMember = ""
    drInfAdmin1.Sections(3).Controls(6).DataField = "Ayud1"
    drInfAdmin1.Sections(3).Controls(7).DataMember = ""
    drInfAdmin1.Sections(3).Controls(7).DataField = "Tota1"
    drInfAdmin1.Sections(3).Controls(8).DataMember = ""
    drInfAdmin1.Sections(3).Controls(8).DataField = "Carn2"
    drInfAdmin1.Sections(3).Controls(9).DataMember = ""
    drInfAdmin1.Sections(3).Controls(9).DataField = "Vale2"
    drInfAdmin1.Sections(3).Controls(10).DataMember = ""
    drInfAdmin1.Sections(3).Controls(10).DataField = "Orde2"
    drInfAdmin1.Sections(3).Controls(11).DataMember = ""
    drInfAdmin1.Sections(3).Controls(11).DataField = "Cuot2"
    drInfAdmin1.Sections(3).Controls(12).DataMember = ""
    drInfAdmin1.Sections(3).Controls(12).DataField = "Ayud2"
    drInfAdmin1.Sections(3).Controls(13).DataMember = ""
    drInfAdmin1.Sections(3).Controls(13).DataField = "Tota2"
    drInfAdmin1.Sections(3).Controls(14).DataMember = ""
    drInfAdmin1.Sections(3).Controls(14).DataField = "Carn3"
    drInfAdmin1.Sections(3).Controls(15).DataMember = ""
    drInfAdmin1.Sections(3).Controls(15).DataField = "Vale3"
    drInfAdmin1.Sections(3).Controls(16).DataMember = ""
    drInfAdmin1.Sections(3).Controls(16).DataField = "Orde3"
    drInfAdmin1.Sections(3).Controls(17).DataMember = ""
    drInfAdmin1.Sections(3).Controls(17).DataField = "Cuot3"
    drInfAdmin1.Sections(3).Controls(18).DataMember = ""
    drInfAdmin1.Sections(3).Controls(18).DataField = "Ayud3"
    drInfAdmin1.Sections(3).Controls(19).DataMember = ""
    drInfAdmin1.Sections(3).Controls(19).DataField = "Tota3"
'totales
    drInfAdmin1.Sections(5).Controls(1).DataMember = ""
    drInfAdmin1.Sections(5).Controls(1).DataField = "Carn2"
    drInfAdmin1.Sections(5).Controls(2).DataMember = ""
    drInfAdmin1.Sections(5).Controls(2).DataField = "Vale2"
    drInfAdmin1.Sections(5).Controls(3).DataMember = ""
    drInfAdmin1.Sections(5).Controls(3).DataField = "Orde2"
    drInfAdmin1.Sections(5).Controls(4).DataMember = ""
    drInfAdmin1.Sections(5).Controls(4).DataField = "Cuot2"
    drInfAdmin1.Sections(5).Controls(5).DataMember = ""
    drInfAdmin1.Sections(5).Controls(5).DataField = "Ayud2"
    drInfAdmin1.Sections(5).Controls(6).DataMember = ""
    drInfAdmin1.Sections(5).Controls(6).DataField = "Tota2"

    drInfAdmin1.Refresh
    Screen.MousePointer = vbDefault
    
    drInfAdmin1.Show

  
  
  Exit Sub

miFinal:
    cOrd.msTermina
    If adoC.State = adStateOpen Then adoC.Close
    Set adoC = Nothing
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
End Sub

Private Sub SumaTodo_InfAdm1(nPrm As Byte)
    'nprm=1 o 2
    Dim sfecha As String
    Dim sAno As String
    '==============
    With cOrd.adoM2
    '==============
    Label5.Caption = "Clasificando totales..."
    Label5.Refresh
    .MoveFirst
    Do While Not cOrd.adoM2.EOF
            'Busca en tabla virtual la fecha: AMM
            If Year(!vencim) < 2000 Then
                sAno = "00"
            Else        'menor que 2000
                sAno = CStr(Year(!vencim) - 2000)
                sAno = IIf(Len(sAno) = 1, "0" & sAno, sAno)
            End If
            sfecha = sAno & IIf(Month(!vencim) > 9, "", "0") & Month(!vencim)
            
            If adoD.RecordCount < 1 Then        'el primer registro
                adoD.AddNew
                adoD("mes") = sfecha
                adoD("fecha") = "10/" & Right(sfecha, 2) & "/" & Left(sfecha, 2)
            Else
                adoD.MoveFirst
                adoD.Find ("mes =" & sfecha)
                If adoD.EOF Then
                    adoD.AddNew
                    adoD("mes") = sfecha
                    adoD("fecha") = "10/" & Right(sfecha, 2) & "/" & Left(sfecha, 2)
                End If
            End If
             'Agrega los totales a la tabla virtual
            If !NoOrden = 1 Then              'cuota
                adoD("cuot" & nPrm) = adoD("cuot" & nPrm) + !valorp
            ElseIf !NoOrden = 2 Then          'ayuda
                 adoD("ayud" & nPrm) = adoD("ayud" & nPrm) + !valorp
            Else
                If cOrd.adoM2!comercio = 190 Then       'vales
                    adoD("vale" & nPrm) = adoD("vale" & nPrm) + !valorp
                ElseIf cOrd.adoM2!comercio = 115 Then     'carnic
                    adoD("carn" & nPrm) = adoD("carn" & nPrm) + !valorp
                Else                                    'credito
                    adoD("orde" & nPrm) = adoD("orde" & nPrm) + !valorp
                End If
            End If
            adoD("Tota" & nPrm) = adoD("Tota" & nPrm) + !valorp
        .MoveNext
   Loop
   '=======
   End With
   '=======
    
End Sub



'=======================================================
Private Sub mInfoGastos(nPrm As Byte)
'=======================================================
    Dim sM1 As String
    Dim sM2 As String
    
    Screen.MousePointer = vbHourglass
    PBar.Visible = False
    Label5.Caption = "Calculando..."
    
    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String
    If nPrm = 1 Then        'entradas y sal funcionarios
        sM = "select *,0.00 as saldo FROM tbl_Gastos INNER JOIN tbl_GastosRubros" & _
                " ON tbl_GastosRubros.sRubro = tbl_Gastos.sRubro" & _
                " WHERE sfecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#" & _
                " ORDER BY sFecha,sHora;"
    Else                    'entradas y sal adm
        sM = "select *,0.00 as saldo FROM tbl_GastosAdm INNER JOIN tbl_GastosRubros" & _
                " ON tbl_GastosRubros.sRubro = tbl_GastosAdm.sRubro" & _
                " WHERE sfecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#" & _
                   " ORDER BY sFecha,sHora;"
    End If
    adoE.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Set adoD = adoE
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
  
    adoD.Sort = "sFecha"
    
    ' Muestra el informe de datos
    'un solo dia
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfGastos.Caption = "Resumen de Gastos del día: " & Text1.Text
        drInfGastos.Title = "Resumen de Gastos " & vbCrLf & "del día: " & Text1.Text
     'varios dias
    Else
        'borra los saldos iniciales, excepto el primero
        Set adoC.ActiveConnection = adoConn
        Set adoC = New ADODB.Recordset
        adoC.Fields.Append "sFecha", adDate
        adoC.Fields.Append "sRubro", adInteger, 2
        adoC.Fields.Append "sDetRubro", adChar, 30
        adoC.Fields.Append "sEntrada", adDouble
        adoC.Fields.Append "sSalida", adDouble
        adoC.Fields.Append "saldo", adDouble
        adoC.Fields.Append "sDetalle", adChar, 200
    
        adoC.CursorType = adOpenDynamic
        adoC.LockType = adLockOptimistic
        adoC.Open
        
        adoD.MoveFirst
        Dim bCuenta As Boolean
        bCuenta = True
        Do While Not adoD.EOF
            If adoD!sRubro = 0 Then
                If bCuenta Then
                    AgregaRegistroGasto
                    bCuenta = False
                End If
            Else
                AgregaRegistroGasto
            End If
            adoD.MoveNext
        Loop
        'Set fjMome.DataGrid1.DataSource = adoC
        'fjMome.Show
        'MsgBox "hola"

        adoD.Close
        Set adoD = adoC
        'Set fjMome.DataGrid1.DataSource = adoD
        'fjMome.Show
        'MsgBox "hola"

        drInfGastos.Caption = "Resumen de Gastos desde: " & Text1.Text & "  hasta: " & Text2.Text
        drInfGastos.Title = "Resumen de Gastos" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
    End If
    'coloca el saldo en el campo saldo
    Dim dSaldo As Double
    dSaldo = 0
    adoD.MoveFirst
    Do While Not adoD.EOF
        dSaldo = dSaldo + adoD!sEntrada - adoD!sSalida
        adoD!saldo = dSaldo
        adoD.MoveNext
    Loop
    
    
    Set drInfGastos.DataSource = adoD
    drInfGastos.DataMember = ""
    
    drInfGastos.Sections(3).Controls(1).DataMember = ""
    drInfGastos.Sections(3).Controls(1).DataField = "sFecha"
    drInfGastos.Sections(3).Controls(2).DataMember = ""
    drInfGastos.Sections(3).Controls(2).DataField = "sRubro"
    drInfGastos.Sections(3).Controls(3).DataMember = ""
    drInfGastos.Sections(3).Controls(3).DataField = "sDetRubro"
    drInfGastos.Sections(3).Controls(4).DataMember = ""
    drInfGastos.Sections(3).Controls(4).DataField = "sEntrada"
    drInfGastos.Sections(3).Controls(6).DataMember = ""
    drInfGastos.Sections(3).Controls(6).DataField = "sSalida"
    drInfGastos.Sections(3).Controls(7).DataMember = ""
    drInfGastos.Sections(3).Controls(7).DataField = "saldo"
    drInfGastos.Sections(3).Controls(5).DataMember = ""
    drInfGastos.Sections(3).Controls(5).DataField = "sDetalle"
    'totales
    drInfGastos.Sections(5).Controls(1).DataMember = ""
    drInfGastos.Sections(5).Controls(1).DataField = "sEntrada"
    drInfGastos.Sections(5).Controls(3).DataMember = ""
    drInfGastos.Sections(5).Controls(3).DataField = "sSalida"
    'drInfGastos.Sections(5).Controls(4).DataMember = ""
    'drInfGastos.Sections(5).Controls(4).DataField = "saldo"
    
    drInfGastos.Refresh
    Screen.MousePointer = vbDefault
    Label5.Caption = ""
    
    drInfGastos.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
    If adoE.State = adStateOpen Then adoE.Close
    Set adoE = Nothing
    Unload Me
End Sub


Private Sub AgregaRegistroGasto()
        adoC.AddNew
        adoC!sfecha = adoD!sfecha
        adoC!sRubro = adoD!sRubro
        adoC!sDetRubro = adoD!sDetRubro
        adoC!sDetalle = adoD!sDetalle
        adoC!sEntrada = adoD!sEntrada
        adoC!sSalida = adoD!sSalida
        adoC!saldo = adoE!saldo
        adoC.Update
End Sub

'=======================================================
Private Sub mInfoCobros3()
'=======================================================
    Dim sM1 As String
    Dim sM2 As String
    
    Screen.MousePointer = vbHourglass
    PBar.Visible = False
    Label5.Caption = "Espere..."
    Label5.Refresh

    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String

    sM = "select *, cdbl(0) as dCarn, cdbl(0) as dVale, cdbl(0) as dCuot, cdbl(0) as dAyu, cdbl(0) as dOrden  FROM tbl_Pagos" & _
            " WHERE not pag_Motivo = 4 AND" & _
            " NOT pag_Motivo = 7 AND" & _
            " pag_Fecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
    adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
    
    'coloca los datos
    Label5.Caption = "Completando..."
    Label5.Refresh
    adoD.MoveFirst
    Do While Not adoD.EOF
        If adoD!pag_NroOrden = 1 Then
            adoD!dcuot = adoD!pag_Valor
        ElseIf adoD!pag_NroOrden = 2 Then
            adoD!dayu = adoD!pag_Valor
        Else
            If adoD!pag_NroCom = 115 Then
                adoD!dcarn = adoD!pag_Valor
           ElseIf adoD!pag_NroCom = 190 Then
                adoD!dvale = adoD!pag_Valor
            Else
                adoD!dorden = adoD!pag_Valor
           End If
        End If
        adoD.MoveNext
    Loop
    
    
    adoD.Sort = "pag_Fecha, pag_NroSoc"
    'adoD.Filter = "pag_motivo <> 7"     'sin recargo
    
    ' Muestra el informe de datos
    'un solo dia
    Label5.Caption = "Armando..."
    Label5.Refresh
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfCobros3.Caption = "Resumen de Cobros del día: " & Text1.Text
        drInfCobros3.Title = "Resumen de Cobros " & vbCrLf & "del día: " & Text1.Text
     'varios dias
    Else
        drInfCobros3.Caption = "Resumen de Cobros desde: " & Text1.Text & "  hasta: " & Text2.Text
        drInfCobros3.Title = "Resumen de Cobros" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
    End If
    
    Set drInfCobros3.DataSource = adoD
    drInfCobros3.DataMember = ""
    
    drInfCobros3.Sections(3).Controls(1).DataMember = ""
    drInfCobros3.Sections(3).Controls(1).DataField = "pag_Fecha"
    drInfCobros3.Sections(3).Controls(2).DataMember = ""
    drInfCobros3.Sections(3).Controls(2).DataField = "pag_NroORden"
    drInfCobros3.Sections(3).Controls(3).DataMember = ""
    drInfCobros3.Sections(3).Controls(3).DataField = "pag_NroSoc"
    drInfCobros3.Sections(3).Controls(4).DataMember = ""
    drInfCobros3.Sections(3).Controls(4).DataField = "dcarn"
    drInfCobros3.Sections(3).Controls(8).DataMember = ""
    drInfCobros3.Sections(3).Controls(8).DataField = "dvale"
    drInfCobros3.Sections(3).Controls(6).DataMember = ""
    drInfCobros3.Sections(3).Controls(6).DataField = "dayu"
    drInfCobros3.Sections(3).Controls(7).DataMember = ""
    drInfCobros3.Sections(3).Controls(7).DataField = "dorden"
    drInfCobros3.Sections(3).Controls(5).DataMember = ""
    drInfCobros3.Sections(3).Controls(5).DataField = "dcuot"
    drInfCobros3.Sections(3).Controls(9).DataMember = ""
    drInfCobros3.Sections(3).Controls(9).DataField = "pag_valor"
    'totales
    drInfCobros3.Sections(5).Controls(1).DataMember = ""
    drInfCobros3.Sections(5).Controls(1).DataField = "dcarn"
    drInfCobros3.Sections(5).Controls(2).DataMember = ""
    drInfCobros3.Sections(5).Controls(2).DataField = "dvale"
    drInfCobros3.Sections(5).Controls(4).DataMember = ""
    drInfCobros3.Sections(5).Controls(4).DataField = "dcuot"
    drInfCobros3.Sections(5).Controls(5).DataMember = ""
    drInfCobros3.Sections(5).Controls(5).DataField = "dayu"
    drInfCobros3.Sections(5).Controls(6).DataMember = ""
    drInfCobros3.Sections(5).Controls(6).DataField = "dorden"
    drInfCobros3.Sections(5).Controls(7).DataMember = ""
    drInfCobros3.Sections(5).Controls(7).DataField = "pag_valor"
    
    drInfCobros3.Refresh
    Screen.MousePointer = vbDefault
    Label5.Caption = ""
    Label5.Refresh

    drInfCobros3.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
    Unload Me
End Sub


'=======================================================
Private Sub mInfoRecargos()
'=======================================================
    Dim sM1 As String
    Dim sM2 As String
    
    Screen.MousePointer = vbHourglass
    PBar.Visible = False
    Label5.Caption = "Espere..."
    Label5.Refresh

    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String
    sM = "select *,space(25) as apellido, space(25) as nombre FROM tbl_Pagos" & _
            " WHERE not pag_Motivo = 4 AND" & _
            " pag_Motivo = 7 AND" & _
            " pag_Fecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
    adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
    
    'coloca los nombres
    Label5.Caption = "Nombres..."
    Label5.Refresh
    cSocio.mfAbreTablaSociosOrdenSocio
    adoD.MoveFirst
    Do While Not adoD.EOF
        cSocio.vlNroSoc = adoD!pag_NroSoc
        If Not cSocio.mfBuscaSocio Then
            adoD!Apellido = "NO ENCONTRADO"
            adoD!nombre = ""
        Else
            adoD!Apellido = cSocio.vsApellido
            adoD!nombre = cSocio.vsNombre
            adoD.Update
        End If
        adoD.MoveNext
    Loop
    
    
    adoD.Sort = "pag_Fecha, pag_NroOrden"
    'adoD.Filter = "pag_motivo <> 7"     'sin recargo

   
    ' Muestra el informe de datos
    'un solo dia
    Label5.Caption = "Armando..."
    Label5.Refresh
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfCobros.Caption = "Resumen de Recargos del día: " & Text1.Text
        drInfCobros.Title = "Resumen de Recargos " & vbCrLf & "del día: " & Text1.Text
     'varios dias
    Else
        drInfCobros.Caption = "Resumen de Recargos desde: " & Text1.Text & "  hasta: " & Text2.Text
        drInfCobros.Title = "Resumen de Recargos" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text
    End If
    
    Set drInfCobros.DataSource = adoD
    drInfCobros.DataMember = ""
    
    drInfCobros.Sections(3).Controls(1).DataMember = ""
    drInfCobros.Sections(3).Controls(1).DataField = "pag_Fecha"
    drInfCobros.Sections(3).Controls(2).DataMember = ""
    drInfCobros.Sections(3).Controls(2).DataField = "pag_NroORden"
    drInfCobros.Sections(3).Controls(3).DataMember = ""
    drInfCobros.Sections(3).Controls(3).DataField = "pag_NroSoc"
    drInfCobros.Sections(3).Controls(4).DataMember = ""
    drInfCobros.Sections(3).Controls(4).DataField = "Apellido"
    drInfCobros.Sections(3).Controls(5).DataMember = ""
    drInfCobros.Sections(3).Controls(5).DataField = "pag_valor"
    drInfCobros.Sections(3).Controls(6).DataMember = ""
    drInfCobros.Sections(3).Controls(6).DataField = "pag_valME"
    drInfCobros.Sections(3).Controls(7).DataMember = ""
    drInfCobros.Sections(3).Controls(7).DataField = "pag_NroCom"
    drInfCobros.Sections(3).Controls(8).DataMember = ""
    drInfCobros.Sections(3).Controls(8).DataField = "NOMBRE"
    drInfCobros.Sections(3).Controls(9).DataMember = ""
    drInfCobros.Sections(3).Controls(9).DataField = "pag_det"
    drInfCobros.Sections(3).Controls(10).DataMember = ""
    drInfCobros.Sections(3).Controls(10).DataField = "pag_NroPago"
   'totales
    drInfCobros.Sections(5).Controls(1).DataMember = ""
    drInfCobros.Sections(5).Controls(1).DataField = "pag_valor"
    drInfCobros.Sections(5).Controls(2).DataMember = ""
    drInfCobros.Sections(5).Controls(2).DataField = "pag_valME"
    
    drInfCobros.Refresh
    Screen.MousePointer = vbDefault
    Label5.Caption = ""
    Label5.Refresh

    drInfCobros.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
    Unload Me
End Sub




'=======================================================
Private Sub mInfoCobrosPorComercio()
'=======================================================
    Dim sM1 As String
    Dim sM2 As String
    
    Screen.MousePointer = vbHourglass
    PBar.Visible = False
    Label5.Caption = "Espere..."
    Label5.Refresh

    
    'Numero de comercio
    '==================
    Dim lCom As Long
    Dim tCom As String
 
    lCom = CLng(Text3.Text)
    tCom = cCom.BuscaComercio(lCom)
    Set cCom = Nothing
    
    'Fecha
    '======
    sM1 = mfInvierteMes(Text1.Text)
    sM2 = mfInvierteMes(Text2.Text)

    'abre la tabla
    '==============
    If adoD.State = adStateOpen Then adoD.Close
    Dim sM As String
    'NI pag_Motivo=4 (anuladas) NI pag_motivo=7 (recargos generados)
    'con No comercio
    'y entre fechas
    sM = "select *,space(25) as apellido, space(25) as nombre FROM tbl_Pagos" & _
            " WHERE not pag_Motivo = 4 AND" & _
            " NOT pag_Motivo = 7 AND" & _
            " pag_NroCom =" & CStr(lCom) & " AND" & _
            " pag_Fecha BETWEEN #" & sM1 & "# AND #" & sM2 & "#;"
    adoD.Open sM, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If adoD.RecordCount < 1 Then
        MsgBox "Sin Registros"
        adoD.Close
        Set adoD = Nothing
        Exit Sub
    End If
    
    'coloca los nombres
    Label5.Caption = "Nombres..."
    Label5.Refresh
    cSocio.mfAbreTablaSociosOrdenSocio
    adoD.MoveFirst
    Do While Not adoD.EOF
        cSocio.vlNroSoc = adoD!pag_NroSoc
        If Not cSocio.mfBuscaSocio Then
            adoD!Apellido = "NO ENCONTRADO"
            adoD!nombre = ""
        Else
            adoD!Apellido = cSocio.vsApellido
            adoD!nombre = cSocio.vsNombre
            adoD.Update
        End If
        adoD.MoveNext
    Loop
    
    
    adoD.Sort = "pag_Fecha, pag_NroOrden"
    'adoD.Filter = "pag_motivo <> 7"     'sin recargo

   
    ' Muestra el informe de datos
    'un solo dia
    Label5.Caption = "Armando..."
    Label5.Refresh
    If CDate(Text1.Text) = CDate(Text2.Text) Then
        drInfCobros.Caption = "Resumen de Cobros del día: " & Text1.Text & " Comercio: " & lCom & "  " & tCom
        drInfCobros.Title = "Resumen de Cobros " & vbCrLf & "del día: " & Text1.Text & vbCrLf & " Comercio: " & lCom & "  " & tCom
     'varios dias
    Else
        drInfCobros.Caption = "Resumen de Cobros desde: " & Text1.Text & "  hasta: " & Text2.Text & " Comercio: " & lCom & "  " & tCom
        drInfCobros.Title = "Resumen de Cobros" & vbCrLf & "desde: " & Text1.Text & "  hasta: " & Text2.Text & vbCrLf & " Comercio: " & lCom & "  " & tCom
    End If
    
    Set drInfCobros.DataSource = adoD
    drInfCobros.DataMember = ""
    
    drInfCobros.Sections(3).Controls(1).DataMember = ""
    drInfCobros.Sections(3).Controls(1).DataField = "pag_Fecha"
    drInfCobros.Sections(3).Controls(2).DataMember = ""
    drInfCobros.Sections(3).Controls(2).DataField = "pag_NroORden"
    drInfCobros.Sections(3).Controls(3).DataMember = ""
    drInfCobros.Sections(3).Controls(3).DataField = "pag_NroSoc"
    drInfCobros.Sections(3).Controls(4).DataMember = ""
    drInfCobros.Sections(3).Controls(4).DataField = "Apellido"
    drInfCobros.Sections(3).Controls(5).DataMember = ""
    drInfCobros.Sections(3).Controls(5).DataField = "pag_valor"
    drInfCobros.Sections(3).Controls(6).DataMember = ""
    drInfCobros.Sections(3).Controls(6).DataField = "pag_valME"
    drInfCobros.Sections(3).Controls(7).DataMember = ""
    drInfCobros.Sections(3).Controls(7).DataField = "pag_NroCom"
    drInfCobros.Sections(3).Controls(8).DataMember = ""
    drInfCobros.Sections(3).Controls(8).DataField = "NOMBRE"
    drInfCobros.Sections(3).Controls(9).DataMember = ""
    drInfCobros.Sections(3).Controls(9).DataField = "pag_det"
    drInfCobros.Sections(3).Controls(10).DataMember = ""
    drInfCobros.Sections(3).Controls(10).DataField = "pag_NroPago"
   'totales
    drInfCobros.Sections(5).Controls(1).DataMember = ""
    drInfCobros.Sections(5).Controls(1).DataField = "pag_valor"
    drInfCobros.Sections(5).Controls(2).DataMember = ""
    drInfCobros.Sections(5).Controls(2).DataField = "pag_valME"
    
    drInfCobros.Refresh
    Screen.MousePointer = vbDefault
    Label5.Caption = ""
    Label5.Refresh

    drInfCobros.Show
    If adoD.State = adStateOpen Then adoD.Close
    Set adoD = Nothing
    Unload Me
End Sub


