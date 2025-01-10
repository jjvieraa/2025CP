VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmodclientes 
   Caption         =   "Modificar Socio"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7845
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9480
      TabIndex        =   39
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdmodsoc 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   7800
      TabIndex        =   37
      Top             =   6840
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   960
      TabIndex        =   25
      Top             =   480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10821
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
      TabPicture(0)   =   "frmmodclientes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "aclara"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "NroCob"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "NroSoc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Fech_ing"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSiguiente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdverdatos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Datos Laborales"
      TabPicture(1)   =   "frmmodclientes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Garantia"
      Tab(1).Control(1)=   "chkAyuda"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "ayuda"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(7)=   "cobrador"
      Tab(1).Control(8)=   "Frame7"
      Tab(1).Control(9)=   "Frame8"
      Tab(1).Control(10)=   "ocupacion"
      Tab(1).Control(11)=   "cmdSiguiente2"
      Tab(1).Control(12)=   "cmdAnterior"
      Tab(1).Control(13)=   "ingresos"
      Tab(1).Control(14)=   "Label7"
      Tab(1).Control(15)=   "Label14"
      Tab(1).Control(16)=   "Label15"
      Tab(1).Control(17)=   "Label16"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Dependientes"
      TabPicture(2)   =   "frmmodclientes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "cmdAnterior2"
      Tab(2).ControlCount=   2
      Begin VB.TextBox Garantia 
         Height          =   285
         Left            =   -73035
         TabIndex        =   70
         ToolTipText     =   "F2=Busca"
         Top             =   840
         Width           =   1305
      End
      Begin VB.CheckBox chkAyuda 
         Caption         =   "Colabora con Ayuda Social"
         Height          =   495
         Left            =   -67320
         TabIndex        =   69
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdverdatos 
         Caption         =   "Ver Datos"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dependientes"
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
         Height          =   4815
         Left            =   -73560
         TabIndex        =   55
         Top             =   960
         Width           =   7365
         Begin VB.CheckBox depautorizado 
            Caption         =   "Autorizado/a a comprar"
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   31
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox deprelacion 
            Height          =   375
            Left            =   1080
            TabIndex        =   30
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox depnombre 
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   26
            Top             =   480
            Width           =   5415
         End
         Begin VB.CommandButton depingresa 
            Caption         =   "Ingresa"
            Height          =   375
            Left            =   360
            TabIndex        =   32
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton depmodifica 
            Caption         =   "Modifica"
            Height          =   375
            Left            =   360
            TabIndex        =   33
            Top             =   3240
            Width           =   1095
         End
         Begin VB.CommandButton depborra 
            Caption         =   "Borra"
            Height          =   375
            Left            =   360
            TabIndex        =   34
            Top             =   3720
            Width           =   1095
         End
         Begin VB.PictureBox DBGrid1 
            Height          =   2295
            Left            =   1560
            ScaleHeight     =   2235
            ScaleWidth      =   4995
            TabIndex        =   56
            Top             =   2280
            Width           =   5055
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Left            =   2640
            TabIndex        =   28
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox depnumero 
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox depci 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   29
            Top             =   960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Caption         =   "C.I."
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   61
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label19 
            Caption         =   "Num."
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "Relación"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Nombre"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   58
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Nacim."
            Height          =   255
            Left            =   2040
            TabIndex        =   57
            Top             =   1080
            Width           =   615
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
         Left            =   1560
         TabIndex        =   46
         Top             =   1800
         Width           =   6615
         Begin VB.ComboBox est_civil 
            DataField       =   "Est_Civil"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   4440
            TabIndex        =   11
            Top             =   3360
            Width           =   1455
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
         Begin VB.TextBox localidad 
            DataField       =   "Localidad"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   2175
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
         Begin VB.TextBox nombre 
            DataField       =   "Nombre"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   2760
            TabIndex        =   6
            Top             =   1080
            Width           =   2415
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
            Left            =   2760
            TabIndex        =   9
            Top             =   2640
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
         Begin MSComCtl2.DTPicker Fech_nac 
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            Top             =   3360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   24707073
            CurrentDate     =   36556
            MaxDate         =   402133
            MinDate         =   214
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   255
            Left            =   3480
            TabIndex        =   54
            Top             =   3420
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfonos"
            Height          =   255
            Left            =   2760
            TabIndex        =   53
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Apellidos"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Nacimiento:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   3420
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Localidad"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Dirección"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Nombres"
            Height          =   255
            Left            =   2880
            TabIndex        =   48
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Cedula de Identidad:"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   420
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Unidad donde presta Servicios"
         Height          =   855
         Left            =   -70440
         TabIndex        =   45
         Top             =   3240
         Width           =   3015
         Begin VB.ComboBox U_Servicio 
            DataField       =   "U_Servicio"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Unidad a la que Pertenece"
         Height          =   855
         Left            =   -73560
         TabIndex        =   44
         Top             =   3240
         Width           =   3015
         Begin VB.ComboBox U_Pertenece 
            DataField       =   "U_Pertenece"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Situación Laboral"
         Height          =   855
         Left            =   -73560
         TabIndex        =   43
         Top             =   2160
         Width           =   3015
         Begin VB.ComboBox cmbstlab 
            Height          =   315
            Left            =   240
            TabIndex        =   16
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
         TabIndex        =   42
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Categoría"
         Height          =   855
         Left            =   -70440
         TabIndex        =   41
         Top             =   2160
         Width           =   3015
         Begin VB.ComboBox cmbcategoria 
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.ComboBox cobrador 
         DataField       =   "Cobrador"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   -68400
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Frame Frame7 
         Caption         =   "Grado"
         Height          =   855
         Left            =   -73560
         TabIndex        =   40
         Top             =   4320
         Width           =   3015
         Begin VB.ComboBox grado 
            DataField       =   "Grado"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   600
            TabIndex        =   20
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         Height          =   855
         Left            =   -70440
         TabIndex        =   36
         Top             =   4320
         Width           =   4095
         Begin VB.CheckBox Cred_Auto 
            Caption         =   "Credito Autorizado"
            DataField       =   "Cred_Auto"
            DataSource      =   "DtClientes"
            Height          =   375
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Limite 
            DataField       =   "Limite"
            DataSource      =   "DtClientes"
            Height          =   315
            Left            =   1920
            TabIndex        =   22
            Top             =   360
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            Caption         =   "Limite de Credito:"
            Height          =   255
            Left            =   600
            TabIndex        =   38
            Top             =   420
            Width           =   1695
         End
      End
      Begin VB.TextBox ocupacion 
         DataField       =   "Ocupacion"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   -72960
         TabIndex        =   14
         Top             =   1500
         Width           =   3495
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Sig&uiente >>"
         Height          =   375
         Left            =   8280
         TabIndex        =   12
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdSiguiente2 
         Caption         =   "Sig&uiente >>"
         Height          =   375
         Left            =   -66720
         TabIndex        =   24
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "<< &Anterior"
         Height          =   375
         Left            =   -74640
         TabIndex        =   23
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnterior2 
         Caption         =   "<< &Anterior"
         Height          =   375
         Left            =   -74760
         TabIndex        =   35
         Top             =   5400
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Fech_ing 
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   37293
         MaxDate         =   402133
         MinDate         =   214
      End
      Begin MSMask.MaskEdBox NroSoc 
         DataField       =   "NroSoc"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   960
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
      Begin MSMask.MaskEdBox NroCob 
         DataField       =   "NroCob"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox ingresos 
         DataField       =   "Ingresos"
         DataSource      =   "DtClientes"
         Height          =   315
         Left            =   -68400
         TabIndex        =   15
         Top             =   1500
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
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "Garantía:"
         Height          =   255
         Left            =   -73965
         TabIndex        =   71
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   6240
         TabIndex        =   68
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Nº de Cobro"
         Height          =   255
         Left            =   3840
         TabIndex        =   67
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nº del Socio"
         Height          =   255
         Left            =   1680
         TabIndex        =   66
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label aclara 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3555
         TabIndex        =   65
         Top             =   8220
         Width           =   1770
      End
      Begin VB.Label Label14 
         Caption         =   "Cobrador"
         Height          =   255
         Left            =   -68400
         TabIndex        =   64
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Ingresos:"
         Height          =   255
         Left            =   -69120
         TabIndex        =   63
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Ocupación:"
         Height          =   255
         Left            =   -73920
         TabIndex        =   62
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmmodclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoclientes As New ADODB.Recordset
Dim adomodificar As New ADODB.Command
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
Private Sub cmdmodsoc_Click()

Dim CodEstCiv As Integer
Dim CodCatSoc As Integer
Dim codunidper As Integer
Dim codpresserv As Integer
Dim CodGrado As Integer
Dim CodSitLab As Integer
Dim i As Integer
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
rstcatsoc.MoveNext
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
        codunidper = rstupert!idupertenece
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
        codpresserv = rstprestserv!iduservicio
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

'modificar
If adoclientes.State = adStateOpen Then adoclientes.Close
Set adoclientes.ActiveConnection = adoconn
adoclientes.Open "select * from TBL_Socios where nrosoc = " & Val(NroSoc.Text) & " ", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

Set adomodificar.ActiveConnection = adoconn
adomodificar.CommandText = "update TBL_Socios set nrocob = '" & NroCob.Text & "', fech_ing = '" & Me.Fech_ing.Value & "', apellido = '" & Me.Apellido.Text & "', nombre = '" & Me.nombre.Text & "', direccion = '" & Me.direccion.Text & "', localidad = '" & Me.localidad.Text & "', tel = '" & Me.Tel.Text & "', fech_nac = '" & Me.Fech_nac.Value & "', codcatsoc = " & CodCatSoc & ", codsitlab = " & CodSitLab & ", CodUnidPer = " & codunidper & ", codpresserv = " & codpresserv & ", codgrado = " & CodGrado & ", ci = '" & Me.ci.Text & "', codestciv = " & CodEstCiv & ", ocupacion = '" & Me.ocupacion.Text & "', ayuda = '" & chkAyuda & "', cobrador = '" & Val(Me.cobrador.Text) & "', ingresos = '" & Me.ingresos.Text & "', cred_auto = '" & Me.Cred_Auto & "', limite = '" & Me.Limite.Text & "' where nrosoc = " & Val(Me.NroSoc.Text) & " "
adomodificar.Execute
MsgBox "Registro Modificado", vbInformation, "Modificaciones"

U_Pertenece.Text = ""
U_Servicio.Text = ""
Tel.Text = ""
ocupacion.Text = ""
NroSoc.Text = ""
NroCob.Text = ""
nombre.Text = ""
Apellido.Text = ""
localidad.Text = ""
ingresos.Text = ""
grado.Text = ""
est_civil.Text = ""
direccion.Text = ""
cmbcategoria.Text = ""
cmbstlab.Text = ""
Limite.Text = ""
cobrador.Text = ""
ci.Text = "_.___.___-_"

SSTab1.Tab = 0
cmdmodsoc.Enabled = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(0) = True

End Sub


Private Sub cmdAnterior_Click()
SSTab1.Tab = 0
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
End Sub

Private Sub cmdAnterior2_Click()
SSTab1.Tab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = False
cmdmodsoc.Enabled = False
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
cmdmodsoc.Enabled = True
End Sub

Private Sub cmdverdatos_Click()
On Error GoTo error
Dim i As Integer

Set adoclientes.ActiveConnection = adoconn
If adoclientes.State = adStateOpen Then adoclientes.Close
    adoclientes.Open "select * from TBL_Socios where NroSoc = " & Val(NroSoc.Text) & "", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
If adoclientes.RecordCount = 0 Then
    MsgBox "No existe el Socio", vbCritical, "Error"
Else
    
    Tel.Text = adoclientes!Tel
    ocupacion.Text = adoclientes!ocupacion
    NroSoc.Text = adoclientes!NroSoc
    NroCob.Text = adoclientes!NroCob
    nombre.Text = adoclientes!nombre
    Apellido.Text = adoclientes!Apellido
    localidad.Text = adoclientes!localidad
    ingresos.Text = adoclientes!ingresos
    direccion.Text = adoclientes!direccion
    cobrador.Text = adoclientes!cobrador
    Fech_ing.Value = adoclientes!Fech_ing
    Fech_nac.Value = adoclientes!Fech_nac
    ci.Text = adoclientes!ci
    Limite.Text = adoclientes!Limite

'////**** búsqueda de descripciones***/////

'categoria
Set rstcatsoc = New Recordset
rstcatsoc.Open "select * from catsocio", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstcatsoc.RecordCount <> 0 Then
        For i = 1 To rstcatsoc.RecordCount
            If adoclientes!CodCatSoc = rstcatsoc!idcatsoc Then
                cmbcategoria.Text = rstcatsoc!Desc
                Exit For
            Else
                rstcatsoc.MoveNext
            End If
        Next i
    Else
        msgtablas
    End If

' grado
Set rstgrado = New Recordset
rstgrado.Open "select * from grado", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstgrado.RecordCount <> 0 Then
        For i = 1 To rstgrado.RecordCount
            If adoclientes!CodGrado = rstgrado!idgrado Then
                grado.Text = rstgrado!Desc
                Exit For
            Else
                rstgrado.MoveNext
            End If
        Next i
    Else
        msgtablas
   End If

' estado civil
Set rstestciv = New Recordset
rstestciv.Open "select * from estcivil", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstestciv.RecordCount <> 0 Then
        For i = 1 To rstestciv.RecordCount
            If adoclientes!CodEstCiv = rstestciv!idestciv Then
                est_civil.Text = rstestciv!Desc
                Exit For
            Else
                rstestciv.MoveNext
            End If
        Next i
    Else
        msgtablas
   End If

' Situacion laboral
Set rstsitlab = New Recordset
rstsitlab.Open "select * from SLaboral", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstsitlab.RecordCount <> 0 Then
        For i = 1 To rstsitlab.RecordCount
            If adoclientes!CodSitLab = rstsitlab!idsitlab Then
                cmbstlab.Text = rstsitlab!Desc
                Exit For
            Else
                rstsitlab.MoveNext
            End If
        Next i
    Else
        msgtablas
    End If

' Unidad Pertenece
Set rstupert = New Recordset
rstupert.Open "select * from unidadpert", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstupert.RecordCount <> 0 Then
        For i = 1 To rstupert.RecordCount
            If adoclientes!codunidper = rstupert!idupertenece Then
                U_Pertenece.Text = rstupert!Desc
                Exit For
            Else
                rstupert.MoveNext
            End If
        Next i
    Else
        msgtablas
    End If

' Unidad Servicio
Set rstprestserv = New Recordset
rstprestserv.Open "select * from unidadserv", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstprestserv.RecordCount <> 0 Then
        For i = 1 To rstprestserv.RecordCount
            If adoclientes!codpresserv = rstprestserv!iduservicio Then
                U_Servicio.Text = rstprestserv!Desc
                Exit For
            Else
                rstprestserv.MoveNext
            End If
        Next i
    Else
        msgtablas
    End If
 
If adoclientes!Cred_Auto = 0 Then
    Cred_Auto.Value = vbUnchecked
Else
    Cred_Auto.Value = vbChecked
End If
If adoclientes!ayuda = 0 Then
    chkAyuda.Value = vbUnchecked
Else
    chkAyuda.Value = vbChecked
End If
Exit Sub
End If

error:
    MsgBox Err.Description, vbQuestion, "Mensaje"
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
cmdmodsoc.Enabled = False
U_Pertenece.Text = ""
U_Servicio.Text = ""
Tel.Text = ""
ocupacion.Text = ""
NroSoc.Text = ""
NroCob.Text = ""
nombre.Text = ""
Apellido.Text = ""
localidad.Text = ""
ingresos.Text = ""
grado.Text = ""
est_civil.Text = ""
direccion.Text = ""
cobrador.Text = ""
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
        cmbcategoria.AddItem (rstcatsoc!Desc)
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
        U_Pertenece.AddItem (rstupert!Desc)
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
        U_Servicio.AddItem (rstprestserv!Desc)
        rstprestserv.MoveNext
    Next i
Else
    msgtablas
End If
    cobrador.AddItem 1
    cobrador.AddItem 2
    cobrador.AddItem 3
    cobrador.AddItem 4
    cobrador.AddItem 5
    cobrador.AddItem 6

End Sub

