VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOrdenes 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   840
      Left            =   90
      TabIndex        =   19
      Top             =   4620
      Visible         =   0   'False
      Width           =   7050
   End
   Begin VB.Frame Frame2 
      Caption         =   "Empresa"
      Height          =   630
      Left            =   120
      TabIndex        =   16
      Top             =   3795
      Width           =   7065
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   17
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   5
         Left            =   1485
         TabIndex        =   18
         Top             =   255
         Width           =   4770
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salir"
      Height          =   315
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5565
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1395
      Left            =   105
      TabIndex        =   8
      Top             =   2325
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   2461
      _Version        =   393216
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2130
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   2
         ToolTipText     =   "F2=Socio F3=Cobro"
         Top             =   195
         Width           =   1230
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Disponible:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1785
         Width           =   960
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenes:"
         Height          =   225
         Left            =   135
         TabIndex        =   14
         Top             =   1425
         Width           =   930
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   255
         Index           =   4
         Left            =   1350
         TabIndex        =   13
         Top             =   1755
         Width           =   915
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Top             =   1395
         Width           =   915
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   270
         Index           =   2
         Left            =   1320
         TabIndex        =   11
         Top             =   1065
         Width           =   2010
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   345
         Index           =   1
         Left            =   870
         TabIndex        =   10
         Top             =   705
         Width           =   3540
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Saldo:"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0E0FF&
         Caption         =   "lbl"
         Height          =   345
         Index           =   0
         Left            =   3285
         TabIndex        =   5
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "No.Cobro:"
         Height          =   345
         Left            =   2445
         TabIndex        =   4
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nombre:"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nro.Socio:"
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.Label lblMensajes 
      BackColor       =   &H00C0E0FF&
      Caption         =   "12345678901234567890123456789012345678901234567890"
      Height          =   225
      Left            =   1785
      TabIndex        =   7
      Top             =   5610
      Width           =   5115
   End
End
Attribute VB_Name = "frmOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSocio As New clsSocios
Dim cOrden As New clsOrdenes


Const kNroCob = 0
Const kNombre = 1
Const kSaldoSueldo = 2
Const kOrdenes = 3
Const kDisponible = 4

Const kNroSoc = 0



Private Sub Command1_Click()
Unload Me
End Sub

'=====================================================================
Private Sub Form_Load()
'=====================================================================
InicializaEtiquetas
End Sub

'=====================================================================
Private Sub InicializaEtiquetas()
'=====================================================================
    lbl(kNroCob).Caption = ""
    lbl(kNombre).Caption = ""
    lbl(kSaldoSueldo).Caption = ""
    lbl(kOrdenes).Caption = ""
    lbl(kDisponible).Caption = ""
    lblMensajes.Caption = ""
    Set DataGrid1.DataSource = Nothing
    DataGrid1.Visible = False
End Sub
'=====================================================================
Private Sub Form_Unload(Cancel As Integer)
'=====================================================================
Set cSocio = Nothing
End Sub

'=====================================================================
Private Sub txt_GotFocus(Index As Integer)
'=====================================================================
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index).Text)
End Sub


'=====================================================================
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'=====================================================================
    If KeyCode = 113 Then   'F2
        vpMuestraTabla = kMuestra3Socios
        frmMuestraTabla.Show
    ElseIf KeyCode = 114 Then   'F3
        vpMuestraTabla = kMuestra4Socios    'Por No Cobro
        frmMuestraTabla.Show
    End If
End Sub

'=====================================================================
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'=====================================================================
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

'=====================================================================
Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'=====================================================================
    Select Case Index
        Case kNroSocio
            Call EjecutaOrden

    End Select
End Sub
'=====================================================================
Private Sub EjecutaOrden()
'=====================================================================
Call InicializaEtiquetas
If Val(txt(kNroSoc).Text) = 0 Then      'POR SI NO se digito el No SOCIO
    Exit Sub
End If
lblMensajes.Caption = "Buscando Datos..."
lblMensajes.Refresh
cSocio.fAbreTablaSociosOrdenSocio
cSocio.vlNroSoc = txt(kNroSoc).Text

'BUSCA AL SOCIO
If Not cSocio.fBuscaSocio Then
    MsgBox "4552: Socio no encontrado"
    Exit Sub
End If

'MUESTRA DATOS DEL SOCIO
lbl(kNombre).Caption = Left(cSocio.vsApellido & " " & cSocio.vsNombre, 40)
lbl(kNroCob).Caption = cSocio.vsNroCob
lbl(kSaldoSueldo).Caption = cSocio.vnsLimite

'BUSCA LAS ORDENES QUE TIENE EL SOCIO
cOrden.vlNroSoc = txt(0)
lblMensajes.Caption = "Espere: Buscando Ordenes..."
lblMensajes.Refresh
If Not cOrden.fBuscaOrdenesUnSocio Then     'HUBO PROBLEMAS
    MsgBox "4553: Problemas al Buscar Ordenes"
    Exit Sub
End If
If cOrden.adoMome.RecordCount = 0 Then
    MsgBox "4554: No tiene Ordenes"
    Exit Sub
End If

'MUESTRA LAS ORDENES EN EL DATAGRID
Set DataGrid1.DataSource = cOrden.adoMome
DataGrid1.Columns(0).Visible = False
PoneTitulos
DataGrid1.Refresh
DataGrid1.Visible = True

'CALCULA LO QUE DEBE PAGAR EL PROXIMO MES
lblMensajes.Caption = "Espere: Calculando Ordenes..."
lblMensajes.Refresh
lbl(kOrdenes).Caption = cOrden.fCalcOrdenesAPagarDelSocio



lblMensajes.Caption = ""
End Sub

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
