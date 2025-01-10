VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form fjIngresos2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Socios"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8745
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   71
      Top             =   5580
      Width           =   5535
      Begin VB.CommandButton cmdincluir 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Incluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdgravar 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAlterar 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Modif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Excluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H0080C0FF&
         Caption         =   "Canc"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0080C0FF&
         Caption         =   "Salir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdultimo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Picture         =   "frmingresos2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdprimeiro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Picture         =   "frmingresos2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Picture         =   "frmingresos2.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdProximo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Picture         =   "frmingresos2.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   240
         Width           =   375
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5310
      Left            =   240
      TabIndex        =   70
      Top             =   180
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   9366
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Socios"
      TabPicture(0)   =   "frmingresos2.frx":0C28
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "aclara"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "NroCob"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "NroSoc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSiguiente"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "COP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Fech_ing"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Datos Laborales"
      TabPicture(1)   =   "frmingresos2.frx":0C44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblGarantia"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ayuda"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cobrador"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "ocupacion"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdSiguiente2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkAyuda"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Garantia"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "ingresos"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Dependientes"
      TabPicture(2)   =   "frmingresos2.frx":0C60
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdGuardarDep"
      Tab(2).Control(1)=   "cmdDepIngresa"
      Tab(2).Control(2)=   "DataGrid1"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(4)=   "cmdModifica"
      Tab(2).Control(5)=   "cmdFinModif"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdFinModif 
         Caption         =   "Termina Modif."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68280
         MaskColor       =   &H00808080&
         TabIndex        =   83
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdModifica 
         Caption         =   "Modifica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68280
         MaskColor       =   &H00808080&
         TabIndex        =   82
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox ingresos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69240
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Fech_ing 
         Height          =   285
         Left            =   4800
         TabIndex        =   3
         Top             =   765
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox COP 
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Top             =   765
         Width           =   495
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2220
         Left            =   -74745
         TabIndex        =   62
         Top             =   585
         Visible         =   0   'False
         Width           =   6345
         Begin MSMask.MaskEdBox mskDepFechNac 
            Height          =   315
            Left            =   2640
            TabIndex        =   26
            Top             =   540
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtdepnum 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   915
            TabIndex        =   32
            Top             =   540
            Width           =   735
         End
         Begin VB.CheckBox chkDepAuto 
            Caption         =   "Autorizado/a a comprar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4035
            TabIndex        =   29
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtDepNom 
            Height          =   315
            Left            =   945
            TabIndex        =   25
            Top             =   15
            Width           =   5175
         End
         Begin VB.TextBox deprelacion 
            Height          =   315
            Left            =   930
            TabIndex        =   28
            Top             =   1080
            Width           =   2655
         End
         Begin MSMask.MaskEdBox DepLimite 
            Height          =   375
            Left            =   930
            TabIndex        =   30
            Top             =   1590
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   15
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtdepci 
            Height          =   315
            Left            =   4695
            TabIndex        =   27
            Top             =   540
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
            Mask            =   "#.###.###-#"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            Caption         =   "Nacim."
            Height          =   255
            Left            =   1890
            TabIndex        =   68
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Límite:"
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   1590
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   66
            Top             =   30
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "Relación"
            Height          =   255
            Left            =   0
            TabIndex        =   65
            Top             =   1110
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Num."
            Height          =   255
            Left            =   0
            TabIndex        =   64
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "C.I."
            Height          =   255
            Index           =   0
            Left            =   4245
            TabIndex        =   63
            Top             =   600
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1995
         Left            =   -74760
         TabIndex        =   61
         Top             =   3135
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3519
         _Version        =   393216
         AllowUpdate     =   0   'False
         DefColWidth     =   27
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "depnum"
            Caption         =   "No."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "depci"
            Caption         =   "Documento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "DepNom"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "DepFechNac"
            Caption         =   "F.Nacim."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "deprel"
            Caption         =   "Relación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "DepAuto"
            Caption         =   "Autoriz"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "DepLimite"
            Caption         =   "Límite"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14346
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDepIngresa 
         Caption         =   " Agrega Dep."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68280
         MaskColor       =   &H00808080&
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuardarDep 
         Caption         =   "Guardar Dep."
         Height          =   390
         Left            =   -68295
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Garantia 
         Height          =   285
         Left            =   -73830
         TabIndex        =   12
         ToolTipText     =   "F2=Muestra"
         Top             =   630
         Width           =   1335
      End
      Begin VB.CheckBox chkAyuda 
         Caption         =   " Ayuda Social"
         Height          =   375
         Left            =   -70200
         TabIndex        =   13
         Top             =   540
         Width           =   1815
      End
      Begin VB.CommandButton cmdSiguiente2 
         Caption         =   "Sig&uiente >>"
         Height          =   375
         Left            =   -66720
         TabIndex        =   35
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Sig&uiente >>"
         Height          =   375
         Left            =   8280
         TabIndex        =   34
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox ocupacion 
         DataField       =   "Ocupacion"
         DataSource      =   "DtClientes"
         Height          =   285
         Left            =   -73800
         TabIndex        =   15
         Top             =   1095
         Width           =   3495
      End
      Begin VB.Frame Frame8 
         Height          =   855
         Left            =   -71685
         TabIndex        =   55
         Top             =   3660
         Width           =   4095
         Begin VB.TextBox Limite 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox Cred_Auto 
            Caption         =   "Credito Autorizado"
            DataField       =   "Cred_Auto"
            DataSource      =   "DtClientes"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label21 
            Caption         =   "Limite de Credito:"
            Height          =   255
            Left            =   600
            TabIndex        =   56
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Grado"
         Height          =   855
         Left            =   -74850
         TabIndex        =   54
         Top             =   3615
         Width           =   3015
         Begin VB.ComboBox grado 
            DataField       =   "Grado"
            DataSource      =   "DtClientes"
            Height          =   315
            ItemData        =   "frmingresos2.frx":0C7C
            Left            =   1320
            List            =   "frmingresos2.frx":0C7E
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.ComboBox cobrador 
         DataField       =   "Cobrador"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   -67440
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Categoría"
         Height          =   855
         Left            =   -71715
         TabIndex        =   52
         Top             =   1545
         Width           =   3015
         Begin VB.ComboBox cmbcategoria 
            Height          =   315
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.CheckBox ayuda 
         Caption         =   "Colabora con Ayuda Social"
         DataField       =   "Ayuda"
         DataSource      =   "DtClientes"
         Height          =   495
         Left            =   -68880
         TabIndex        =   51
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Situación Laboral"
         Height          =   855
         Left            =   -74850
         TabIndex        =   50
         Top             =   1545
         Width           =   3015
         Begin VB.ComboBox cmbstlab 
            Height          =   315
            Left            =   255
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Unidad a la que Pertenece"
         Height          =   855
         Left            =   -74865
         TabIndex        =   49
         Top             =   2640
         Width           =   3015
         Begin VB.ComboBox U_Pertenece 
            DataField       =   "U_Pertenece"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   375
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Unidad donde presta Servicios"
         Height          =   855
         Left            =   -71640
         TabIndex        =   33
         Top             =   2640
         Width           =   3015
         Begin VB.ComboBox U_Servicio 
            DataField       =   "U_Servicio"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos Personales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   4095
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   6195
         Begin MSMask.MaskEdBox Fech_nac 
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.TextBox nombre 
            DataField       =   "Nombre"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   2760
            TabIndex        =   6
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox direccion 
            DataField       =   "Direccion"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   240
            TabIndex        =   7
            Top             =   1800
            Width           =   4935
         End
         Begin VB.TextBox localidad 
            DataField       =   "Localidad"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox Apellido 
            DataField       =   "Apellido"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   2295
         End
         Begin VB.ComboBox est_civil 
            DataField       =   "Est_Civil"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   4425
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   3360
            Width           =   1455
         End
         Begin MSMask.MaskEdBox ci 
            DataField       =   "CI"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   1920
            TabIndex        =   4
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox tel 
            DataField       =   "Tel"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   2775
            TabIndex        =   9
            Top             =   2625
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "Cedula de Identidad:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   420
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Nombres"
            Height          =   255
            Left            =   2880
            TabIndex        =   43
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Dirección"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Localidad"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Nacimiento:"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   3420
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Apellidos"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfonos"
            Height          =   255
            Left            =   2760
            TabIndex        =   38
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   255
            Left            =   3480
            TabIndex        =   37
            Top             =   3420
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox NroSoc 
         DataField       =   "NroSoc"
         DataSource      =   "DtClientes"
         Height          =   285
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "F2=Nro F3=Alfab"
         Top             =   765
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox NroCob 
         DataField       =   "NroCob"
         DataSource      =   "DtClientes"
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   765
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "COP"
         Height          =   270
         Left            =   3885
         TabIndex        =   69
         Top             =   525
         Width           =   465
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Left            =   6525
         Top             =   585
         Width           =   1665
      End
      Begin VB.Label lblGarantia 
         Height          =   495
         Left            =   -72360
         TabIndex        =   60
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Garantía:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Ocupación:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Ingresos:"
         Height          =   255
         Left            =   -70215
         TabIndex        =   57
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Cobrador"
         Height          =   255
         Left            =   -68280
         TabIndex        =   53
         Top             =   660
         Width           =   735
      End
      Begin VB.Label aclara 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3555
         TabIndex        =   48
         Top             =   8220
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Nº del Socio"
         Height          =   255
         Left            =   270
         TabIndex        =   47
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Nº de Cobro"
         Height          =   255
         Left            =   2100
         TabIndex        =   46
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   4845
         TabIndex        =   45
         Top             =   540
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adocat 
      Height          =   330
      Left            =   240
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=jimmy"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "jimmy"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CatSocio"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adoestciv 
      Height          =   330
      Left            =   240
      Top             =   6840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=jimmy"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "jimmy"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EstCivil"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "fjIngresos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoSocios As New ADODB.Recordset
Dim adoDep As New ADODB.Recordset
Dim adoMome As New ADODB.Recordset
Dim adoIns As New ADODB.Command

Dim funcao As String











'1=======================================================================================================
Private Sub Form_Load()
'=======================================================================================================

SSTab1.Tab = 0
'1) abre la tabla
Set adoSocios.ActiveConnection = adoconn
If adoSocios.State = adStateOpen Then adoSocios.Close
adoSocios.Open "select * from TBL_Socios ORDER BY NroSOc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

'2) carga grupos en los combobox
Call CargarTablasComboBox

'3) Determina propiedades de los campos
Call PropiedadesDeCampos
   
'4)Inicio
Call inicio
Call cmdPrimeiro_Click

Call BloqueaTexto
 
End Sub 'Form_Load

'2=======================================================================================================
Private Sub CargarTablasComboBox()
'=======================================================================================================
    Dim i As Integer
    
    'Carga Numeros de Cobradores
    cobrador.AddItem 1
    cobrador.AddItem 2
    cobrador.AddItem 3
    cobrador.AddItem 4
    cobrador.AddItem 5
    cobrador.AddItem 6
   cobrador.AddItem 7
   cobrador.AddItem 8
   cobrador.AddItem 9
   cobrador.AddItem 10
   cobrador.AddItem 11
   cobrador.AddItem 12
   cobrador.AddItem 13
   cobrador.AddItem 14
   cobrador.AddItem 15
   cobrador.AddItem 16
   cobrador.AddItem 17
   cobrador.AddItem 18
   cobrador.AddItem 19
   cobrador.AddItem 20
    
    'Cargar Estado Civil
    Set adoMome = New Recordset
    adoMome.Open "select * from EstCivil", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 1 To adoMome.RecordCount
        est_civil.AddItem (adoMome!Desc)
        adoMome.MoveNext
    Next i
       
    'Cargar Categoria
    adoMome.Close
    adoMome.Open "select * from catsocio", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 1 To adoMome.RecordCount
        Me.cmbcategoria.AddItem (adoMome!Desc)
        adoMome.MoveNext
    Next i
    
    'Cargar grado
    adoMome.Close
    adoMome.Open "select * from Grado", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 1 To adoMome.RecordCount
            grado.AddItem (adoMome!Desc)
            adoMome.MoveNext
        Next i
 
    'cargar sitlab
    adoMome.Close
    adoMome.Open "select * from slaboral", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 1 To adoMome.RecordCount
        cmbstlab.AddItem (adoMome!Desc)
        adoMome.MoveNext
    Next i

    'cargar unidad pertenece
    adoMome.Close
    adoMome.Open "select * from unidadpert", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 1 To adoMome.RecordCount
        Me.U_Pertenece.AddItem (adoMome!Desc)
        adoMome.MoveNext
    Next i

    'cargar unidad servicio
    adoMome.Close
    adoMome.Open "select * from unidadserv", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 1 To adoMome.RecordCount
            Me.U_Servicio.AddItem (adoMome!Desc)
            adoMome.MoveNext
        Next i
    adoMome.Close
    Set adoMome = Nothing
End Sub

Private Sub BloqueaTexto()
    Fech_ing.Enabled = False
    Fech_nac.Enabled = False
    ci.Enabled = False
    
    NroSoc.Enabled = True
    NroCob.Enabled = False
    COP.Enabled = False
    Tel.Enabled = False
    ocupacion.Enabled = False
    nombre.Enabled = False
    Apellido.Enabled = False
    localidad.Enabled = False
    ingresos.Enabled = False
    direccion.Enabled = False
    cobrador.Enabled = False
    Limite.Enabled = False
    Garantia.Enabled = False
    
    est_civil.Enabled = False
    cobrador.Enabled = False
    cmbstlab.Enabled = False
    cmbcategoria.Enabled = False
    U_Pertenece.Enabled = False
    U_Servicio.Enabled = False
    grado.Enabled = False

    chkAyuda.Enabled = False
    Cred_Auto.Enabled = False
End Sub

Private Sub DesbloqueaTexto()
    Fech_ing.Enabled = True
    Fech_nac.Enabled = True
    ci.Enabled = True
    
    NroSoc.Enabled = False
    NroCob.Enabled = True
    COP.Enabled = True
    Tel.Enabled = True
    ocupacion.Enabled = True
    nombre.Enabled = True
    Apellido.Enabled = True
    localidad.Enabled = True
    ingresos.Enabled = True
    direccion.Enabled = True
    cobrador.Enabled = True
    Limite.Enabled = True
    Garantia.Enabled = True
    
    est_civil.Enabled = True
    cobrador.Enabled = True
    cmbstlab.Enabled = True
    cmbcategoria.Enabled = True
    U_Pertenece.Enabled = True
    U_Servicio.Enabled = True
    grado.Enabled = True

    chkAyuda.Enabled = True
    Cred_Auto.Enabled = True
End Sub

'3=================================
Private Sub PropiedadesDeCampos()
'==================================
NroSoc.MaxLength = 5
NroCob.MaxLength = 22
Apellido.MaxLength = 25
nombre.MaxLength = 25
direccion.MaxLength = 50
localidad.MaxLength = 30
Tel.MaxLength = 25
ocupacion.MaxLength = 15
COP.MaxLength = 2
End Sub





'4==================================
Private Sub inicio()
'==================================
   adoSocios.Requery
    If adoSocios.RecordCount > 0 Then
        If funcao <> "ALT" Then
            adoSocios.MoveFirst
            ActualizaFormulario
        End If
    Else
        LimpiaBoxes
    End If
    
    cmdincluir.Enabled = True
    cmdSair.Enabled = True
    cmdgravar.Enabled = False
    cmdCancelar.Enabled = False
    
    
    If adoSocios.RecordCount = 0 Then
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
        cmdAnterior.Enabled = False
        cmdProximo.Enabled = False
        cmdprimeiro.Enabled = False
        cmdultimo.Enabled = False
    Else
        cmdAlterar.Enabled = True
        cmdExcluir.Enabled = True
        cmdAnterior.Enabled = True
        cmdProximo.Enabled = True
        cmdprimeiro.Enabled = True
        cmdultimo.Enabled = True
    End If
    

End Sub


'5=======================================================================================================
Private Sub LimpiaBoxes()
'=======================================================================================================
    Fech_ing.Mask = ""
    Fech_ing.Text = ""
    Fech_ing.Mask = "##/##/####"
    Fech_nac.Mask = ""
    Fech_nac.Text = ""
    Fech_nac.Mask = "##/##/####"
    ci.Mask = ""
    ci.Text = ""
    ci.Mask = "#.###.###-#"
    
    NroSoc.Text = ""
    NroCob.Text = ""
    COP.Text = ""
    Tel.Text = ""
    ocupacion.Text = ""
    nombre.Text = ""
    Apellido.Text = ""
    localidad.Text = ""
    ingresos.Text = "0"
    direccion.Text = ""
    cobrador.Text = ""
    Limite.Text = "0"
    Garantia.Text = "0"
    
    est_civil.ListIndex = 0
    cobrador.ListIndex = 0
    cmbstlab.ListIndex = 0
    cmbcategoria.ListIndex = 0
    U_Pertenece.ListIndex = 0
    U_Servicio.ListIndex = 0
    grado.ListIndex = 0
End Sub



'6=======================================================
Private Sub ActualizaFormulario()
'=======================================================
 
NroSoc.Text = "" & adoSocios("nrosoc")
NroCob.Text = "" & adoSocios("nrocob")
Fech_ing.Text = "" & adoSocios("fech_ing")
Apellido.Text = "" & adoSocios("apellido")
nombre.Text = "" & adoSocios("nombre")
direccion.Text = "" & adoSocios("direccion")
localidad.Text = "" & adoSocios("localidad")
Tel.Text = "" & adoSocios("tel")
ci.Text = "" & adoSocios("ci")
Fech_nac.Mask = ""
Fech_nac.Text = Format(adoSocios("Fech_nac"), "short date")
COP.Text = "" & adoSocios("cop")
ocupacion.Text = "" & adoSocios("ocupacion")
ingresos.Text = Format(0 + adoSocios("ingresos"), "#,#0")
Limite.Text = Format(0 + adoSocios("limite"), "#,#0")
Garantia.Text = adoSocios("garantia")


Cred_Auto.Value = IIf(adoSocios("cred_auto"), vbChecked, vbUnchecked)
chkAyuda.Value = IIf(adoSocios("ayuda"), vbChecked, vbUnchecked)

cmbcategoria.ListIndex = adoSocios("CodCatSoc")
cmbstlab.ListIndex = adoSocios("CodSitLab")
U_Pertenece.ListIndex = adoSocios("codunidper")
U_Servicio.ListIndex = adoSocios("codpresserv")
grado.ListIndex = adoSocios("CodGrado")
est_civil.ListIndex = adoSocios("codestciv")
cobrador.ListIndex = adoSocios("cobrador")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If adoSocios.State = adStateOpen Then
    adoSocios.Close
End If
Set adoSocios = Nothing

If adoDep.State = adStateOpen Then adoDep.Close
Set adoDep = Nothing

Set fjIngresos2 = Nothing
End Sub



'======================================================
'VALIDACIONES =========================================

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
End Sub




Private Sub Apellido_LostFocus()
Apellido.Text = UCase(Apellido.Text)
End Sub





Private Sub localidad_LostFocus()
localidad.Text = UCase(localidad.Text)
End Sub


Private Sub nombre_LostFocus()
nombre.Text = UCase(nombre.Text)
End Sub

Private Sub direccion_lostfocus()
direccion.Text = UCase(direccion.Text)
End Sub



Private Sub NroCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 115 And funcao = "INC" Then   'F4 PARA INGRESAR UN NUMERO DETERMINADO
        NroSoc.Enabled = True
        NroSoc.SetFocus
    End If

End Sub

Private Sub ocupacion_Validate(Cancel As Boolean)
ocupacion.Text = UCase(ocupacion.Text)
End Sub
Private Sub est_civil_LostFocus()
SSTab1.Tab = 1
Garantia.SetFocus
End Sub



Private Sub chkDepAuto_Click()
If Cred_Auto.Value = vbChecked Then
    Limite.Enabled = True
Else
    Limite.Enabled = False
End If
End Sub


'=======================================================================================================
Private Sub Garantia_KeyDown(KeyCode As Integer, Shift As Integer)  'jv
'=======================================================================================================
    If KeyCode = 113 Then   'F2
        vpMuestraTabla = kMstrSocAlf
        fjMuestraTabla.Show
    End If
End Sub


Private Sub Garantia_LostFocus()        'jv
'=======================================================================================================
Call MuestraNombreGarantia
End Sub

'=======================================================================================================
Private Sub MuestraNombreGarantia()     'jv
'=======================================================================================================
    Dim sCriterio As String
    On Error GoTo E0101
    
    'No tiene garantia
    If Not IsNumeric(Garantia.Text) Then Exit Sub
    If CInt(Garantia.Text) = 0 Then
        lblGarantia.Caption = ""
        Exit Sub
    End If
    'Busca el numero de socio
    Set adoMome = New ADODB.Recordset
    Set adoMome.ActiveConnection = adoconn
    If adoMome.State = adStateOpen Then adoMome.Close
    adoMome.Open "SELECT NroSoc,Apellido,Nombre FROM TBL_Socios ORDER by NroSoc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    sCriterio = "NroSoc =" & CInt(Garantia.Text)
    adoMome.MoveFirst
    adoMome.Find (sCriterio)
    If Not adoMome.EOF Then
        lblGarantia.Caption = adoMome!Apellido & " " & adoMome!nombre
    Else
        lblGarantia.Caption = "Desconocido."
    End If
    adoMome.Close
    Set adoMome = Nothing
    Exit Sub

E0101:
    MsgBox ("ERROR 13231: " & Err.Description)

End Sub
'=======================================================================================================
Private Sub NroSoc_KeyDown(KeyCode As Integer, Shift As Integer)        'JV
'=======================================================================================================
    If KeyCode = 113 Then   'F2
        vpMuestraTabla = kMstrSoc1
        fjMuestraTabla.Show
    ElseIf KeyCode = 114 Then 'f3
        vpMuestraTabla = kMstrSocAlf5
        fjMuestraTabla.Show
    End If

End Sub


Private Sub NroSoc_LostFocus()
Dim nM As Integer


'VERIFICA QUE NO EXISTA EL SOCIO en un ingreso
If funcao = "INC" Then 'SI ES UN CLIENTE NUEVO
    adoMome.Open "select * from TBL_Socios where nrosoc = " & CLng(Me.NroSoc.Text) & " ", adoconn, adOpenKeyset, adLockOptimistic, adCmdText ''
    nM = adoMome.RecordCount
    adoMome.Close
    Set adoMome = Nothing
    If nM > 0 Then
        MsgBox "El socio ya existe", vbCritical, "Registro Duplicado"
        NroSoc.SetFocus
        Exit Sub
    End If
Else
    'busca el socio:
    adoSocios.MoveFirst
    adoSocios.Find ("NroSoc =" & CLng(NroSoc.Text))
    If adoSocios.EOF Then adoSocios.MoveLast
    ActualizaFormulario
End If
End Sub

Private Sub NroSoc_Validate(Cancel As Boolean)
If NroSoc.Text = "" Or _
    IsNull(NroSoc.Text) Or _
    Not IsNumeric(NroSoc.Text) Then
        Cancel = True
End If
End Sub



'D E P E N D I E N T E S=================

'DEP 1 ==========================================
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then
    funcao = "DEP"
    Frame9.Enabled = False
    Frame9.Visible = False
    cmdGuardarDep.Enabled = False
    cmdFinModif.Visible = False
    Call BuscayMuestraDepend
Else
    If funcao = "DEP" Then
        DepCierra
        Frame9.Enabled = True
        Frame9.Visible = True
    End If
End If
End Sub


'DEP 2=======================================================================================================
Private Sub BuscayMuestraDepend()           'jv
'=======================================================================================================
    Dim sCriterio As String
    On Error GoTo E0102
    
    'EL CAMPO No SOCIO ESTA VACIO
    If Not IsNumeric(Me.NroSoc.Text) Then Exit Sub
    
    'BUSCA LOS DEPENDIENTES
    Set adoDep = New ADODB.Recordset
    If adoDep.State = adStateOpen Then adoDep.Close
    adoDep.Open "select depnum,depci,DepNom,DepFechNac,deprel," & _
        "DepAuto,DepLimite FROM TBL_Dependientes where nrosoc = " & _
        CLng(Me.NroSoc.Text) & ";", _
        adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    DataGrid1.Visible = False
    Set DataGrid1.DataSource = adoDep
    DataGrid1.Columns(0).Width = 400
    DataGrid1.Columns(3).Width = 1500
    'DataGrid1.Columns(6).DataFormat
    DataGrid1.Visible = True
    Exit Sub

E0102:
    MsgBox ("ERROR 13232: " & Err.Description & " " & Err.Number)
End Sub

'DEP 3=======================================================================================================
Private Sub cmdDepIngresa_Click()
'va a agregar un nuevo dependiente
'=======================================================================================================

'busca el numero de dependiente
txtdepnum.Text = adoDep.RecordCount + 1
Frame3.Visible = True
cmdGuardarDep.Enabled = True
cmdDepIngresa.Enabled = False
cmdModifica.Enabled = False
Me.txtDepNom.SetFocus
End Sub


'DEP 4=======================================================================================================
Private Sub cmdGuardarDep_Click()
'=======================================================================================================
On Error GoTo E03344
        'guarda los datos
        
        Set adoIns.ActiveConnection = adoconn
        adoIns.CommandText = "insert into TBL_Dependientes values(" & _
            CLng(NroSoc.Text) & ", " & _
            CLng(txtdepnum.Text) & ", '" & _
            txtdepci.Text & "', '" & _
            UCase(txtDepNom.Text) & "', '" & _
            mskDepFechNac.Text & "', '" & _
            UCase(deprelacion.Text) & "', '" & _
            chkDepAuto & "'," & _
            CLng(DepLimite.Text) & ",'" & _
            vpnFuncionario & "','" & _
            Date & "','" & _
            Time & "')"
        adoIns.Execute
        'cera las variables
        Call InicializaCamposDepend
        'muestra la nueva situacion
        Call BuscayMuestraDepend
        Frame3.Visible = False
        cmdGuardarDep.Enabled = False
        cmdDepIngresa.Enabled = True
        cmdModifica.Enabled = True
        Set adoIns = Nothing
        Exit Sub
E03344:
    MsgBox "ERROR 03344: " & Err.Description & " " & Err.Number
End Sub





'DEP 5=======================================================================================================
Private Sub InicializaCamposDepend()
'=======================================================================================================
        deprelacion.Text = ""
        txtdepci.Mask = ""
        txtdepci.Text = ""
        txtdepci.Mask = "#.###.###-#"
        mskDepFechNac.Mask = ""
        mskDepFechNac.Text = ""
        mskDepFechNac.Mask = "##/##/####"
        txtDepNom.Text = ""
        txtdepnum.Text = ""
        DepLimite.Text = ""
        DepLimite.Enabled = False
        chkDepAuto.Value = vbUnchecked

End Sub


'DEP 6==============================
Private Sub DepCierra()
    If adoDep.State = adStateOpen Then
        adoDep.Close
    End If
    Set adoDep = Nothing
End Sub


'dep 7 ============================
Private Sub cmdModifica_Click()
cmdModifica.Enabled = False
cmdModifica.Visible = False
cmdDepIngresa.Enabled = False
cmdFinModif.Enabled = True
cmdFinModif.Visible = True
DataGrid1.AllowDelete = True
DataGrid1.AllowUpdate = True
End Sub

'dep 8 ============================
Private Sub cmdFinModif_Click()
adoDep.Update
cmdModifica.Enabled = True
cmdModifica.Visible = True
cmdDepIngresa.Enabled = True
cmdFinModif.Enabled = False
cmdFinModif.Visible = False
DataGrid1.AllowDelete = False
DataGrid1.AllowUpdate = False
End Sub




'B O T O N E S ============================================






Private Sub ActualizaCampos()   'Me intriga el nombre ??

adoSocios("fech_ing") = CDate(Fech_ing.Text)
adoSocios("Fech_nac") = CDate(Fech_nac.Text)

adoSocios("ingresos") = CSng(ingresos.Text)
adoSocios("limite") = CSng(Limite.Text)
adoSocios("garantia") = CInt(Garantia.Text)

adoSocios("nrosoc") = CLng(NroSoc.Text)
adoSocios("nrocob") = NroCob.Text
adoSocios("apellido") = Apellido.Text
adoSocios("nombre") = nombre.Text
adoSocios("direccion") = direccion.Text
adoSocios("localidad") = localidad.Text
adoSocios("tel") = Tel.Text
adoSocios("ci") = ci.Text
adoSocios("cop") = COP.Text
adoSocios("ocupacion") = ocupacion.Text

adoSocios("cred_auto") = Cred_Auto
adoSocios("ayuda") = chkAyuda

adoSocios("cobrador") = cobrador.ListIndex

adoSocios("CodCatSoc") = cmbcategoria.ListIndex
adoSocios("CodSitLab") = cmbstlab.ListIndex
adoSocios("codunidper") = U_Pertenece.ListIndex
adoSocios("codpresserv") = U_Servicio.ListIndex
adoSocios("CodGrado") = grado.ListIndex
adoSocios("codestciv") = est_civil.ListIndex


End Sub

'================================
Private Function ValidaCampos() As Boolean
'================================
If NroCob.Text = "" Or Not IsNumeric(NroCob.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    NroCob.SetFocus
    ValidaCampos = False
    Exit Function
End If
If Apellido.Text = "" Or IsNull(Apellido.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    Apellido.SetFocus
    ValidaCampos = False
    Exit Function
End If
If nombre.Text = "" Or IsNull(nombre.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    nombre.SetFocus
    ValidaCampos = False
    Exit Function
End If
If ci.Text = "" Or IsNull(ci.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    ci.SetFocus
    ValidaCampos = False
    Exit Function
End If
If Not IsDate(Fech_nac.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    Fech_nac.SetFocus
    ValidaCampos = False
    Exit Function
End If
ValidaCampos = True
End Function


'================================
Private Sub cmdgravar_Click()
'================================
    If Not ValidaCampos Then
        Exit Sub
    End If
    If funcao = "INC" Then
        adoSocios.AddNew
        ActualizaCampos
        adoSocios.Update
        inicio
    End If
    If funcao = "ALT" Then
        ActualizaCampos
        adoSocios.Update
        'MsgBox rsMsg("msg3"), vbOKOnly, "Aviso"
        
        'para localizar o registro
        inicio
    End If
    BloqueaTexto
    SSTab1.TabEnabled(2) = True
End Sub


'================================
Private Sub cmdincluir_Click()
'================================
    'incluir un socio
    funcao = "INC"
    SSTab1.Tab = 0
    SSTab1.TabEnabled(2) = False
    Botones
    LimpiaBoxes
    DesbloqueaTexto
    'NroSoc.Enabled = True
    'toma el ultimo No para aumentarlo
    adoSocios.MoveLast
    NroSoc.Text = adoSocios("nrosoc") + 1
    Fech_ing.Mask = ""
    Fech_ing.Text = ""
    Fech_ing.Text = Format(Date, "short date")
    NroCob.SetFocus
End Sub

'================================
Private Sub cmdSair_Click()
'================================
   Unload Me

End Sub


Private Sub cmdAlterar_Click()
    funcao = "ALT"
    DesbloqueaTexto
    SSTab1.Tab = 0
    SSTab1.TabEnabled(2) = False
    Botones
    ActualizaFormulario
End Sub


Private Sub cmdCancelar_Click()
    funcao = ""
    BloqueaTexto
    SSTab1.TabEnabled(2) = True
    If adoSocios.RecordCount > 0 Then
        adoSocios.MoveFirst
        ActualizaFormulario
    Else
        LimpiaBoxes
    End If
    inicio

End Sub

Private Sub cmdExcluir_Click()
    funcao = "EXC"
    If MsgBox("Confirma ?", vbYesNo, "Confirmando !") = vbYes Then
            adoSocios.Delete
            If adoSocios.RecordCount > 0 Then
            ' Mostra o registro anterior pois esse nao existe mais
                cmdAnterior_Click
            Else
                LimpiaBoxes
            End If
    End If
    inicio
    BloqueaTexto
End Sub


'================================
Private Sub cmdPrimeiro_Click()
'================================
    If adoSocios.RecordCount > 0 Then
        adoSocios.MoveFirst
        If adoSocios.BOF = True Then
            adoSocios.MoveFirst
        End If
        ActualizaFormulario
    Else
        MsgBox "Sin Registros", 16, "Aviso"
        Exit Sub
    End If
End Sub

'================================
Private Sub cmdProximo_Click()
'================================
    If adoSocios.RecordCount > 0 Then
        adoSocios.MoveNext
        If adoSocios.EOF = True Then
            adoSocios.MovePrevious
        End If
        ActualizaFormulario
    Else
        MsgBox "Sin Registros", 16, "Aviso"
        Exit Sub
    End If
    
End Sub

'================================
Private Sub cmdAnterior_Click()
'================================
    If adoSocios.RecordCount > 0 Then
    adoSocios.MovePrevious
    If adoSocios.BOF = True Then
        adoSocios.MoveNext
    End If
    ActualizaFormulario
    Else
        MsgBox "Sin Registros", 16, "Aviso"
        Exit Sub
    End If
End Sub

'================================
Private Sub cmdultimo_Click()
'================================
    If adoSocios.RecordCount > 0 Then
        adoSocios.MoveLast
        If adoSocios.EOF = True Then
            adoSocios.MoveLast
        End If
        ActualizaFormulario
    Else
        MsgBox "Sin Registros", 16, "Aviso"
        Exit Sub
    End If
End Sub
    
Private Sub Botones()

    cmdincluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdSair.Enabled = True
    cmdgravar.Enabled = True
    cmdCancelar.Enabled = True
    
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = False
    cmdprimeiro.Enabled = False
    cmdultimo.Enabled = False

End Sub


