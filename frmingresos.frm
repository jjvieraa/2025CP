VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fjIngresos 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Socios"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8745
   Begin TabDlg.SSTab SSTab1 
      Height          =   5310
      Left            =   165
      TabIndex        =   73
      Top             =   180
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   9366
      _Version        =   393216
      Tab             =   2
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
      TabPicture(0)   =   "frmingresos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "COP"
      Tab(0).Control(1)=   "cmdSiguiente"
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(3)=   "Fech_ing"
      Tab(0).Control(4)=   "NroSoc"
      Tab(0).Control(5)=   "NroCob"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Image1"
      Tab(0).Control(8)=   "aclara"
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(11)=   "Label3"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Datos Laborales"
      TabPicture(1)   =   "frmingresos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "lblGarantia"
      Tab(1).Control(5)=   "ingresos"
      Tab(1).Control(6)=   "Frame6"
      Tab(1).Control(7)=   "Frame5"
      Tab(1).Control(8)=   "Frame1"
      Tab(1).Control(9)=   "ayuda"
      Tab(1).Control(10)=   "Frame2"
      Tab(1).Control(11)=   "cobrador"
      Tab(1).Control(12)=   "Frame7"
      Tab(1).Control(13)=   "Frame8"
      Tab(1).Control(14)=   "ocupacion"
      Tab(1).Control(15)=   "cmdSiguiente2"
      Tab(1).Control(16)=   "chkAyuda"
      Tab(1).Control(17)=   "Garantia"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Dependientes"
      TabPicture(2)   =   "frmingresos.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdGuardarDep"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdOtroSocio"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmddepingresa"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DataGrid1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.TextBox COP 
         Height          =   285
         Left            =   -71220
         TabIndex        =   2
         Top             =   750
         Width           =   750
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2220
         Left            =   255
         TabIndex        =   65
         Top             =   585
         Width           =   6345
         Begin VB.TextBox txtdepnum 
            Height          =   315
            Left            =   915
            TabIndex        =   26
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   1080
            Width           =   2655
         End
         Begin MSMask.MaskEdBox DepLimite 
            Height          =   375
            Left            =   930
            TabIndex        =   31
            Top             =   1590
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin VB.PictureBox cboDepFchNac 
            Height          =   315
            Left            =   2685
            ScaleHeight     =   255
            ScaleWidth      =   1275
            TabIndex        =   27
            Top             =   540
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtdepci 
            Height          =   315
            Left            =   4695
            TabIndex        =   28
            Top             =   540
            Width           =   1575
            _ExtentX        =   2778
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
            TabIndex        =   71
            Top             =   570
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Límite:"
            Height          =   255
            Left            =   0
            TabIndex        =   70
            Top             =   1590
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   69
            Top             =   30
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "Relación"
            Height          =   255
            Left            =   0
            TabIndex        =   68
            Top             =   1110
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Num."
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   510
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "C.I."
            Height          =   255
            Index           =   0
            Left            =   4245
            TabIndex        =   66
            Top             =   600
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1995
         Left            =   330
         TabIndex        =   64
         Top             =   3135
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3519
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
      Begin VB.CommandButton cmddepingresa 
         Caption         =   " Otro Dependiente"
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
         Left            =   6705
         MaskColor       =   &H00808080&
         TabIndex        =   24
         Top             =   2220
         Width           =   1455
      End
      Begin VB.CommandButton cmdOtroSocio 
         Caption         =   "Otro Socio"
         Height          =   390
         Left            =   6705
         TabIndex        =   35
         Top             =   1665
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuardarDep 
         Caption         =   "Guardar Dep."
         Height          =   390
         Left            =   6705
         TabIndex        =   32
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Garantia 
         Height          =   285
         Left            =   -73830
         TabIndex        =   12
         ToolTipText     =   "F2=Busca"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkAyuda 
         Caption         =   "Colabora con Ayuda Social"
         Height          =   375
         Left            =   -70200
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdSiguiente2 
         Caption         =   "Sig&uiente >>"
         Height          =   375
         Left            =   -66720
         TabIndex        =   37
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Sig&uiente >>"
         Height          =   375
         Left            =   -66720
         TabIndex        =   36
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
         TabIndex        =   58
         Top             =   3660
         Width           =   4095
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
         Begin MSMask.MaskEdBox Limite 
            DataField       =   "Limite"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   1935
            TabIndex        =   23
            Top             =   420
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            Caption         =   "Limite de Credito:"
            Height          =   255
            Left            =   600
            TabIndex        =   59
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Grado"
         Height          =   855
         Left            =   -74850
         TabIndex        =   57
         Top             =   3615
         Width           =   3015
         Begin VB.ComboBox grado 
            DataField       =   "Grado"
            DataSource      =   "DtClientes"
            Height          =   315
            ItemData        =   "frmingresos.frx":0054
            Left            =   1320
            List            =   "frmingresos.frx":0056
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Situación Laboral"
         Height          =   855
         Left            =   -74850
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   34
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
         Left            =   -74865
         TabIndex        =   39
         Top             =   1170
         Width           =   6195
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
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#.###.###-#"
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
         Begin VB.PictureBox Fech_nac 
            Height          =   315
            Left            =   1815
            ScaleHeight     =   255
            ScaleWidth      =   1275
            TabIndex        =   10
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Cedula de Identidad:"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   420
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Nombres"
            Height          =   255
            Left            =   2880
            TabIndex        =   46
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Dirección"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Localidad"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Nacimiento:"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   3420
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Apellidos"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfonos"
            Height          =   255
            Left            =   2760
            TabIndex        =   41
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   255
            Left            =   3480
            TabIndex        =   40
            Top             =   3420
            Width           =   1095
         End
      End
      Begin VB.PictureBox Fech_ing 
         Height          =   285
         Left            =   -70260
         ScaleHeight     =   225
         ScaleWidth      =   1395
         TabIndex        =   3
         Top             =   765
         Width           =   1455
      End
      Begin MSMask.MaskEdBox NroSoc 
         DataField       =   "NroSoc"
         DataSource      =   "DtClientes"
         Height          =   285
         Left            =   -74805
         TabIndex        =   0
         ToolTipText     =   "F2=Muestra"
         Top             =   765
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   -73110
         TabIndex        =   1
         Top             =   750
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox ingresos 
         DataField       =   "Ingresos"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   -69420
         TabIndex        =   16
         Top             =   1050
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "COP"
         Height          =   270
         Left            =   -71115
         TabIndex        =   72
         Top             =   510
         Width           =   465
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Left            =   -68475
         Top             =   585
         Width           =   1665
      End
      Begin VB.Label lblGarantia 
         Height          =   255
         Left            =   -72360
         TabIndex        =   63
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Garantía:"
         Height          =   255
         Left            =   -74775
         TabIndex        =   62
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Ocupación:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Ingresos:"
         Height          =   255
         Left            =   -70215
         TabIndex        =   60
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Cobrador"
         Height          =   255
         Left            =   -68280
         TabIndex        =   56
         Top             =   660
         Width           =   735
      End
      Begin VB.Label aclara 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -71445
         TabIndex        =   51
         Top             =   8220
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Nº del Socio"
         Height          =   255
         Left            =   -74730
         TabIndex        =   50
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Nº de Cobro"
         Height          =   255
         Left            =   -72900
         TabIndex        =   49
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   -70155
         TabIndex        =   48
         Top             =   540
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc adodepe 
      Height          =   330
      Left            =   2400
      Top             =   6360
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
      CommandType     =   1
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
      RecordSource    =   "select * from dependientes, auxdep where nrosoc = nsoc;"
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
   Begin MSAdodcLib.Adodc Adogrado 
      Height          =   330
      Left            =   2880
      Top             =   6000
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
      RecordSource    =   "Grado"
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
   Begin MSAdodcLib.Adodc Adouserv 
      Height          =   330
      Left            =   1695
      Top             =   6000
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
      RecordSource    =   "UnidadServ"
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
   Begin MSAdodcLib.Adodc Adoupert 
      Height          =   330
      Left            =   2880
      Top             =   6360
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
      RecordSource    =   "UnidadPert"
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
   Begin MSAdodcLib.Adodc Adositlab 
      Height          =   330
      Left            =   1680
      Top             =   6360
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
      RecordSource    =   "SLaboral"
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
   Begin VB.CommandButton btnSocioIngresa 
      Caption         =   "&Guardar"
      Height          =   330
      Left            =   5430
      TabIndex        =   33
      Top             =   5595
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   6840
      TabIndex        =   38
      Top             =   5595
      Width           =   1335
   End
End
Attribute VB_Name = "fjIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoSocios As New ADODB.Recordset
Dim adoInsertar As New ADODB.Command
Dim adoDep As New ADODB.Recordset
Dim adoMome As New ADODB.Recordset
Dim WithEvents rstestciv As Recordset
Attribute rstestciv.VB_VarHelpID = -1
Dim WithEvents rstcatsoc As Recordset
Attribute rstcatsoc.VB_VarHelpID = -1
Dim WithEvents rstupert As Recordset
Attribute rstupert.VB_VarHelpID = -1
Dim WithEvents rstprestserv As Recordset
Attribute rstprestserv.VB_VarHelpID = -1
Dim WithEvents rstgrado As Recordset
Attribute rstgrado.VB_VarHelpID = -1
Dim WithEvents rstsitlab As Recordset
Attribute rstsitlab.VB_VarHelpID = -1
Dim kNoSaleAun As Boolean




Private Sub Apellido_LostFocus()
Apellido.Text = UCase(Apellido.Text)
End Sub


'=======================================================================================================
Private Sub Form_Load()
'=======================================================================================================
Dim i As Integer

    SSTab1.Tab = 0
    
    'es un ingreso
    If vpFormMovim = kFormIngresa Then
        SSTab1.TabEnabled(2) = False
        btnSocioIngresa.Enabled = False
        Cred_Auto.Value = vbUnchecked
        Call InicializaCamposSocios
    End If
    
    Call CargarTablasComboBox
    Call CargarValidaciones

 
End Sub 'Form_Load



Private Sub localidad_LostFocus()
localidad.Text = UCase(localidad.Text)
End Sub


Private Sub nombre_LostFocus()
nombre.Text = UCase(nombre.Text)
End Sub

Private Sub ocupacion_Validate(Cancel As Boolean)
ocupacion.Text = UCase(ocupacion.Text)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)    'jv
If SSTab1.Tab = 1 Then btnSocioIngresa.Enabled = True
End Sub



Private Sub est_civil_LostFocus()
SSTab1.Tab = 1
End Sub



'=======================================================================================================
Private Sub InicializaCamposSocios()
'=======================================================================================================
    tel.Text = ""
    ocupacion.Text = ""
    NroSoc.Text = ""
    NroCob.Text = ""
    nombre.Text = ""
    Apellido.Text = ""
    localidad.Text = ""
    ingresos.Text = ""
    'grado.Text = ""
    'est_civil.Text = ""
    direccion.Text = ""
    cobrador.Text = ""
    COP.Text = ""
    ci.Text = "_.___.___-_"
    Limite.Text = ""
End Sub


'=======================================================================================================
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
    
    'Cargar Estado Civil
    Set rstestciv = New Recordset
    rstestciv.Open "select * from EstCivil", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstestciv.RecordCount <> 0 Then
        For i = 1 To rstestciv.RecordCount
            est_civil.AddItem (rstestciv!Desc)
            rstestciv.MoveNext
        Next i
    Else
        msgtablas
    End If
       
    'Cargar Categoria
    Set rstcatsoc = New Recordset
    rstcatsoc.Open "select * from catsocio", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstcatsoc.RecordCount <> 0 Then
        For i = 1 To rstcatsoc.RecordCount
            Me.cmbcategoria.AddItem (rstcatsoc!Desc)
            rstcatsoc.MoveNext
        Next i
    Else
        msgtablas
    End If
    
    'Cargar grado
    Set rstgrado = New Recordset
    rstgrado.Open "select * from Grado", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstgrado.RecordCount <> 0 Then
        For i = 1 To rstgrado.RecordCount
            grado.AddItem (rstgrado!Desc)
            rstgrado.MoveNext
        Next i
    Else
        msgtablas
    End If

    'cargar sitlab
    Set rstsitlab = New Recordset
    rstsitlab.Open "select * from slaboral", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstsitlab.RecordCount <> 0 Then
        For i = 1 To rstsitlab.RecordCount
            cmbstlab.AddItem (rstsitlab!Desc)
            rstsitlab.MoveNext
        Next i
    Else
        msgtablas
    End If

    'cargar unidad pertenece
    Set rstupert = New Recordset
    rstupert.Open "select * from unidadpert", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        If rstupert.RecordCount <> 0 Then
        For i = 1 To rstupert.RecordCount
            Me.U_Pertenece.AddItem (rstupert!Desc)
            rstupert.MoveNext
        Next i
    Else
        msgtablas
    End If

    'cargar unidad servicio
    Set rstprestserv = New Recordset
    rstprestserv.Open "select * from unidadserv", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstprestserv.RecordCount <> 0 Then
        For i = 1 To rstprestserv.RecordCount
            Me.U_Servicio.AddItem (rstprestserv!Desc)
            rstprestserv.MoveNext
        Next i
    Else
        msgtablas
    End If
    

End Sub

'=======================================================================================================
Private Sub CargarValidaciones()
'=======================================================================================================
    'VALIDACIONES      jv
    NroCob.MaxLength = 11
    Apellido.MaxLength = 25
    nombre.MaxLength = 25
    direccion.MaxLength = 50
    localidad.MaxLength = 30
    tel.MaxLength = 25
    COP.MaxLength = 2

End Sub


Private Sub chkDepAuto_Click()
If chkDepAuto.Value = vbChecked Then
    DepLimite.Enabled = True
Else
    DepLimite.Enabled = False
End If
End Sub





'=======================================================================================================
Private Sub cmdGuardarDep_Click()
'=======================================================================================================
On Error GoTo E03344
        cmdGuardarDep.Enabled = False
        'guarda los datos
        Set adoInsertar.ActiveConnection = adoconn
        'adoInsertar.CommandText = "insert into TBL_Dependientes values(" & _
            Val(NroSoc.Text) & ", " & _
            Val(txtdepnum.Text) & ", '" & _
            txtdepci.Text & "', '" & _
            txtDepNom.Text & "', '" & _
            cboDepFchNac.Value & "', '" & _
            deprelacion.Text & "', '" & _
            chkDepAuto & "'," & _
            Val(DepLimite.Text) & ",'" & _
            vpnFuncionario & "','" & _
            Date & "','" & _
            Time & "')"
        adoInsertar.Execute
        'cera las variables
        Call InicializaCamposDepend
        'muestra la nueva situacion
        Call BuscayMuestraDepend
        Exit Sub
E03344:
    MsgBox "ERROR 03344: " & Err.Description & " " & Err.Number
End Sub

'=======================================================================================================
Private Sub InicializaCamposDepend()
'=======================================================================================================
        deprelacion.Text = ""
        txtdepci.Text = "_.___.___-_"
        txtDepNom.Text = ""
        txtdepnum.Text = ""
        DepLimite.Text = ""
        DepLimite.Enabled = False
        chkDepAuto.Value = vbUnchecked

End Sub


'=======================================================================================================
Private Sub cmdOtroSocio_Click()
'=======================================================================================================
        Call InicializaCamposSocios
        
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        btnSocioIngresa.Enabled = False
        SSTab1.Tab = 0
        
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdSiguiente_Click()
SSTab1.Tab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = False
End Sub

Private Sub cmdSiguiente2_Click()
SSTab1.Tab = 2
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = True
btnSocioIngresa.Enabled = True
End Sub

'=======================================================================================================
Private Sub cmdDepIngresa_Click()
'=======================================================================================================

'busca el numero de dependiente
txtdepnum.Text = adoMome.RecordCount + 1
Frame3.Visible = True
cmdGuardarDep.Enabled = True
Me.txtDepNom.SetFocus
End Sub


'=======================================================================================================
Private Sub Garantia_KeyDown(KeyCode As Integer, Shift As Integer)  'jv
'=======================================================================================================
    If KeyCode = 113 Then   'F2
        vpMuestraTabla = kMuestraSocios
        fjMuestraTabla.Show
    End If
End Sub


'=======================================================================================================
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
        lblGarantia.Caption = ""
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
        kNoSaleAun = True     'MOMENTANEMENTE SE VA A OTRA PANT JV
        vpMuestraTabla = kMuestra2Socios
        fjMuestraTabla.Show
    End If

End Sub

Private Sub NroSoc_LostFocus()
Dim i As Integer
Dim res As Integer
Set adoSocios.ActiveConnection = adoconn
If adoSocios.State = adStateOpen Then adoSocios.Close
adoSocios.Open "select * from TBL_Socios where nrosoc = " & Val(Me.NroSoc.Text) & " ", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

If kNoSaleAun = True Then      'jv
    kNoSaleAun = False
    Exit Sub
End If
    
If adoSocios.RecordCount <> 0 Then
    res = MsgBox("El socio ya existe" & Chr(13) & "¿Desea salir de la pantalla?", vbExclamation + vbYesNo, "Registro Duplicado")
    If res = vbNo Then
         NroSoc.SetFocus
    Else
        Unload Me
    End If
End If

If NroSoc.Text = "" Then
    res = MsgBox("Debe Ingresar un Número para continuar" & Chr(13) & "¿Desea salir de la pantalla?", vbExclamation + vbYesNo, "Registro Duplicado")
    If res = vbNo Then
         NroSoc.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Sub btnSocioIngresa_Click()      'Me intriga el nombre ??
Dim CodEstCiv As Integer
Dim CodCatSoc As Integer
Dim codupert As Integer
Dim codprestserv As Integer
Dim CodGrado As Integer
Dim CodSitLab As Integer
Dim i As Integer
Dim res As Integer
'On Error GoTo error

'Buscar código Categoría
rstcatsoc.MoveFirst
For i = 1 To rstcatsoc.RecordCount
    If Trim(UCase(rstcatsoc!Desc)) = Trim(UCase(cmbcategoria.Text)) Then
        CodCatSoc = rstcatsoc!idcatsoc
        Exit For
    End If
rstcatsoc.MoveNext
Next i
If cmbcategoria.Text = "" Then CodCatSoc = 0
'Fin Buscar código categoria

'Buscar código estado civil
rstestciv.MoveFirst
For i = 1 To rstestciv.RecordCount
    If Trim(UCase(rstestciv!Desc)) = Trim(UCase(est_civil.Text)) Then
        CodEstCiv = rstestciv!idestciv
        Exit For
    End If
rstestciv.MoveNext
Next i
If est_civil.Text = "" Then CodEstCiv = 0
'Fin Buscar código estado civil

'Buscar código grado
rstgrado.MoveFirst
For i = 1 To rstgrado.RecordCount
    If Trim(UCase(rstgrado!Desc)) = Trim(UCase(grado.Text)) Then
        CodGrado = rstgrado!idgrado
        Exit For
    End If
rstgrado.MoveNext
Next i
If grado.Text = "" Then CodGrado = 0
'Fin Buscar código grado

'Buscar código Unidad Pertenece
rstupert.MoveFirst
For i = 1 To rstupert.RecordCount
    If Trim(UCase(rstupert!Desc)) = Trim(UCase(U_Pertenece.Text)) Then
        codupert = rstupert!idupertenece
        Exit For
    End If
rstupert.MoveNext
Next i
If U_Pertenece.Text = "" Then codupert = 0
'Fin Buscar código Unidad Pertenece

'Buscar código Unidad Servicio
rstprestserv.MoveFirst
For i = 1 To rstprestserv.RecordCount
    If Trim(UCase(rstprestserv!Desc)) = Trim(UCase(U_Servicio.Text)) Then
        codprestserv = rstprestserv!iduservicio
        Exit For
    End If
rstprestserv.MoveNext
Next i
If U_Servicio.Text = "" Then codprestserv = 0
'Fin Buscar código Unidad Servicio

'Buscar código Situacion Laboral
rstsitlab.MoveFirst
For i = 1 To rstsitlab.RecordCount
    If Trim(UCase(rstsitlab!Desc)) = Trim(UCase(cmbstlab.Text)) Then
        CodSitLab = rstsitlab!idsitlab
        Exit For
    End If
rstsitlab.MoveNext
Next i
If cmbstlab.Text = "" Then CodSitLab = 0
'Fin Buscar código Situacion Laboral

Set adoSocios.ActiveConnection = adoconn
If adoSocios.State = adStateOpen Then adoSocios.Close
adoSocios.Open "select * from TBL_Socios where nrosoc = " & Val(Me.NroSoc.Text) & " ", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

'Si no existe el socio
If adoSocios.RecordCount = 0 Then
    'res = MsgBox("Desea ingresar más dependientes ", vbQuestion + vbYesNo, "Registro Duplicado")
    res = MsgBox("Desea ingresar el socio? ", vbQuestion + vbYesNo, "Registro Duplicado")   'jv
    If res = vbYes Then 'Ingresa el socio
       Set adoInsertar.ActiveConnection = adoconn
        'adoInsertar.CommandText = "insert into TBL_Socios values( " & _
            Val(NroSoc.Text) & ",'" & NroCob.Text & "','" & _
            Fech_ing.Value & "','" & UCase(Apellido.Text) & "','" & _
            UCase(nombre.Text) & "','" & _
            direccion.Text & "','" & _
            UCase(localidad.Text) & "','" & _
            tel.Text & "','" & Fech_nac.Value & "'," & _
            CodCatSoc & ", " & CodSitLab & ", " & _
            codupert & ", " & codprestserv & ", " & _
            CodGrado & ", '" & ci.Text & "', " & _
            CodEstCiv & ", '" & ocupacion.Text & "','" & _
            Me.chkAyuda & "', " & Val(cobrador.Text) & " ," & _
            Val(ingresos.Text) & ",'" & Me.Cred_Auto & "'," & _
            Val(Limite.Text) & "," & Val(Garantia.Text) & ",'" & _
            COP.Text & "','" & vpnFuncionario & "','" & _
            Date & "','" & Time & "')"
       adoInsertar.Execute
       mensaje 'mensaje de OK
    Else                'no lo ingresa
        'If res = vbNo Then      jv
            txtDepNom.SetFocus
        'End If                   jv
    End If
End If




'Botones y Fichas
Me.btnSocioIngresa.Enabled = False
SSTab1.TabEnabled(2) = True

SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(0) = False
cmdGuardarDep.Enabled = False
cmdOtroSocio.Enabled = True
cmddepingresa.Enabled = True
Frame3.Visible = False
SSTab1.Tab = 2

Call BuscayMuestraDepend


Exit Sub

error:
   MsgBox Err.Description

End Sub
'=======================================================================================================
Private Sub BuscayMuestraDepend()           'jv
'=======================================================================================================
    Dim sCriterio As String
    On Error GoTo E0102
    
    'No tiene garantia
    If Not IsNumeric(Me.NroSoc.Text) Then Exit Sub
    
    'Busca el numero de socio
    Set adoMome = New ADODB.Recordset
    Set adoMome.ActiveConnection = adoconn
    If adoMome.State = adStateOpen Then adoMome.Close
    adoMome.Open "select depnum,depci,DepNom,DepFechNac,deprel,DepAuto,DepLimite FROM TBL_Dependientes where nrosoc = " & Val(Me.NroSoc.Text) & ";", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid1.DataSource = adoMome
    'adoMome.Close
    'Set adoMome = Nothing
    Exit Sub

E0102:
    MsgBox ("ERROR 13232: " & Err.Description & " " & Err.Number)
End Sub





