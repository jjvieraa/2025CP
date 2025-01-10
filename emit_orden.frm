VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOrdenes1 
   Caption         =   "Emisión de Ordenes"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   9540
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   35
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Height          =   765
      Left            =   90
      TabIndex        =   30
      Top             =   3585
      Width           =   4845
      Begin VB.TextBox Razon 
         DataField       =   "Razon"
         DataSource      =   "DTComercios"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1575
         TabIndex        =   31
         Top             =   345
         Width           =   3135
      End
      Begin MSMask.MaskEdBox CodComer 
         DataField       =   "Nro"
         DataSource      =   "DTComercios"
         Height          =   315
         Left            =   135
         TabIndex        =   32
         Top             =   375
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Nº del Comercio"
         Height          =   255
         Left            =   135
         TabIndex        =   34
         Top             =   105
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   1575
         TabIndex        =   33
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2235
      Left            =   45
      TabIndex        =   21
      Top             =   90
      Width           =   6315
      Begin MSMask.MaskEdBox disponible 
         Height          =   315
         Left            =   4035
         TabIndex        =   22
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   255
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox numsocio 
         DataField       =   "NroSoc"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   90
         TabIndex        =   23
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox numcobro 
         DataField       =   "NroCob"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   1770
         TabIndex        =   24
         Top             =   405
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CopiaDisp 
         Height          =   315
         Left            =   4050
         TabIndex        =   25
         Top             =   435
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   0
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Nº del Socio"
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Nº de Cobro"
         Height          =   255
         Left            =   1770
         TabIndex        =   28
         Top             =   165
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "D I S P O N I B L E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4050
         TabIndex        =   27
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Nombres"
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   1890
         Width           =   2175
      End
   End
   Begin MSMask.MaskEdBox limite 
      Height          =   285
      Left            =   4080
      TabIndex        =   20
      Top             =   675
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenes Emitidas y Pendientes de Pago"
      Height          =   3525
      Left            =   5010
      TabIndex        =   18
      Top             =   2520
      Width           =   6765
      Begin VB.PictureBox DBGrid1 
         Height          =   3135
         Left            =   1305
         ScaleHeight     =   3075
         ScaleWidth      =   6585
         TabIndex        =   19
         Top             =   495
         Width           =   6645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fotografia"
      Height          =   2325
      Left            =   6540
      TabIndex        =   17
      Top             =   120
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   60
         Top             =   210
         Width           =   1995
      End
   End
   Begin VB.ComboBox CbMes 
      Height          =   315
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4590
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   615
      Left            =   1710
      TabIndex        =   12
      Top             =   5070
      Width           =   3015
      Begin VB.OptionButton Moneda 
         Caption         =   "R$"
         Height          =   255
         Index           =   2
         Left            =   2175
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "U$S"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "$"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox NombComp 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox Mes 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   3270
      TabIndex        =   1
      Top             =   4590
      Width           =   375
   End
   Begin MSMask.MaskEdBox ValOrden 
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   5310
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CantCuot 
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   6030
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ValorCuot 
      Height          =   315
      Left            =   1590
      TabIndex        =   7
      Top             =   6030
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CodComp 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Label Label14 
      Caption         =   "Codigo y Nombre del que realiza la compra"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label11 
      Caption         =   "Valor de la Cuota"
      Height          =   255
      Left            =   1590
      TabIndex        =   8
      Top             =   5790
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Cant. de Cuotas"
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   5790
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Valor de la Orden"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   5070
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Para el Presupuesto"
      Height          =   255
      Left            =   1590
      TabIndex        =   2
      Top             =   4350
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   4350
      Width           =   1335
   End
End
Attribute VB_Name = "frmOrdenes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

