VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ingcomercio 
   Caption         =   "Ingreso de Comercios"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   6900
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdinsertcom 
      Caption         =   "&Guardar"
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
      Left            =   3420
      TabIndex        =   26
      Top             =   6135
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
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
      Left            =   1905
      TabIndex        =   25
      Top             =   6135
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Comercio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   5895
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   6630
      Begin VB.TextBox txtComNom 
         DataField       =   "Razon"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox Rubro 
         DataField       =   "Rubro"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Direc 
         DataField       =   "Direc"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1110
         TabIndex        =   11
         Top             =   2265
         Width           =   5175
      End
      Begin VB.TextBox Razon 
         DataField       =   "Razon"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Afiliación"
         Height          =   855
         Left            =   1110
         TabIndex        =   7
         Top             =   3585
         Width           =   3135
         Begin VB.OptionButton optcoop 
            Caption         =   "Cooperador"
            Height          =   255
            Left            =   1680
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optadherido 
            Caption         =   "Adherido"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox Trab_Coop 
         Caption         =   "Trabaja c/socio Cooperador"
         DataField       =   "Trab_Coop"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   1110
         TabIndex        =   6
         Top             =   4785
         Width           =   1575
      End
      Begin VB.ComboBox Cierre 
         DataField       =   "Cierre"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5070
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox Desc 
         DataField       =   "Desc"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5550
         TabIndex        =   4
         Top             =   4065
         Width           =   735
      End
      Begin VB.CheckBox Discrimina 
         Caption         =   "Discriminar Gastos"
         DataField       =   "Discrimina"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   3030
         TabIndex        =   3
         Top             =   4785
         Width           =   1575
      End
      Begin VB.CheckBox Convenio 
         Caption         =   "Convenio c/cuota mensual fija"
         DataField       =   "Convenio"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   4710
         TabIndex        =   2
         Top             =   4785
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker Fech_ing 
         Height          =   315
         Left            =   4710
         TabIndex        =   1
         Top             =   825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox Tel 
         DataField       =   "Tel"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1110
         TabIndex        =   13
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nro 
         DataField       =   "Nro"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1110
         TabIndex        =   14
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox RUC 
         DataField       =   "RUC"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2910
         TabIndex        =   15
         Top             =   3000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Nº del Comercio"
         Height          =   255
         Left            =   1110
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   1110
         TabIndex        =   23
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Rubro"
         Height          =   255
         Left            =   2550
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   2910
         TabIndex        =   21
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   1110
         TabIndex        =   20
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   4710
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Descuento (%)"
         Height          =   255
         Left            =   5190
         TabIndex        =   17
         Top             =   3705
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Cierre Día"
         Height          =   255
         Left            =   5070
         TabIndex        =   16
         Top             =   2760
         Width           =   855
      End
   End
End
Attribute VB_Name = "ingcomercio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


