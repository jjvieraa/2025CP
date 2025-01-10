VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fjVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Datos"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   10
      Tab             =   8
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Categoría"
      TabPicture(0)   =   "fjVarios.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdGuardar"
      Tab(0).Control(1)=   "txtcat"
      Tab(0).Control(2)=   "cmdBorrar"
      Tab(0).Control(3)=   "DataGrid1"
      Tab(0).Control(4)=   "adocatsocio"
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label4"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Estado Civil"
      TabPicture(1)   =   "fjVarios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "adoEstCiv"
      Tab(1).Control(1)=   "cmdDelEstCivil"
      Tab(1).Control(2)=   "txtEstCiv"
      Tab(1).Control(3)=   "cmdSaveEstCiv"
      Tab(1).Control(4)=   "DGEstCiv"
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(7)=   "Label1"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Grado"
      TabPicture(2)   =   "fjVarios.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "adoGrado"
      Tab(2).Control(1)=   "cmdDelGdo"
      Tab(2).Control(2)=   "txtGdo"
      Tab(2).Control(3)=   "cmdSaveGdo"
      Tab(2).Control(4)=   "DGGdo"
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(7)=   "Label5"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Situación Laboral"
      TabPicture(3)   =   "fjVarios.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(3)=   "DGSitLab"
      Tab(3).Control(4)=   "cmdSaveSitLab"
      Tab(3).Control(5)=   "txtSitLab"
      Tab(3).Control(6)=   "cmdDelSitLab"
      Tab(3).Control(7)=   "adoSLaboral"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Unidad Perteneciente"
      TabPicture(4)   =   "fjVarios.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "adoUniPer"
      Tab(4).Control(1)=   "cmdDelUniPer"
      Tab(4).Control(2)=   "txtUniPer"
      Tab(4).Control(3)=   "cmdSaveUniPer"
      Tab(4).Control(4)=   "DGUniPer"
      Tab(4).Control(5)=   "Frame1"
      Tab(4).Control(6)=   "Label10"
      Tab(4).Control(7)=   "Label9"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Unidad de Servicios"
      TabPicture(5)   =   "fjVarios.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DGUniSer"
      Tab(5).Control(1)=   "adoUniSer"
      Tab(5).Control(2)=   "cmdDelUniSer"
      Tab(5).Control(3)=   "txtUniSer"
      Tab(5).Control(4)=   "cmdSaveUniSer"
      Tab(5).Control(5)=   "Frame6"
      Tab(5).Control(6)=   "Label12"
      Tab(5).Control(7)=   "Label11"
      Tab(5).ControlCount=   8
      TabCaption(6)   =   "Grupos"
      TabPicture(6)   =   "fjVarios.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).Control(1)=   "txtGrupos"
      Tab(6).Control(2)=   "adoGrupo"
      Tab(6).Control(3)=   "DataGrid2"
      Tab(6).Control(4)=   "Label14"
      Tab(6).Control(5)=   "Label13"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Parámetros"
      TabPicture(7)   =   "fjVarios.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "mskUltimo"
      Tab(7).Control(1)=   "txtOrden"
      Tab(7).Control(2)=   "cmdGrabarParam"
      Tab(7).Control(3)=   "cmdSalir(7)"
      Tab(7).Control(4)=   "Label15(1)"
      Tab(7).Control(5)=   "Label15(0)"
      Tab(7).ControlCount=   6
      TabCaption(8)   =   "Funcionarios"
      TabPicture(8)   =   "fjVarios.frx":00E0
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "DataGrid3"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "adoFunc"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "Frame8"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "T.Cambio"
      TabPicture(9)   =   "fjVarios.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "DataGrid4"
      Tab(9).Control(1)=   "adoTC"
      Tab(9).Control(2)=   "Frame9"
      Tab(9).ControlCount=   3
      Begin VB.Frame Frame9 
         Height          =   975
         Left            =   -74640
         TabIndex        =   69
         Top             =   4200
         Width           =   6255
         Begin VB.CommandButton Command3 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   73
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Borrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   72
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   4680
            TabIndex        =   71
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame8 
         Height          =   975
         Left            =   360
         TabIndex        =   63
         Top             =   4320
         Width           =   6255
         Begin VB.CommandButton cmdAgregaFunc 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   4680
            TabIndex        =   66
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdBorraFunc 
            Caption         =   "&Borrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   65
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdGuardarFunc 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   64
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSMask.MaskEdBox mskUltimo 
         Height          =   315
         Left            =   -73320
         TabIndex        =   62
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtOrden 
         Height          =   315
         Left            =   -73320
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabarParam 
         Caption         =   "Guardar"
         Height          =   315
         Left            =   -70800
         TabIndex        =   58
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Index           =   7
         Left            =   -71880
         TabIndex        =   57
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame Frame7 
         Height          =   975
         Left            =   -73560
         TabIndex        =   51
         Top             =   4320
         Width           =   4935
         Begin VB.CommandButton cmdGuardarGrupo 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   56
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdBorrarGrupo 
            Caption         =   "&Borrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   55
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   3360
            TabIndex        =   52
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtGrupos 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -72480
         MaxLength       =   30
         TabIndex        =   49
         Top             =   1440
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DGUniSer 
         Bindings        =   "fjVarios.frx":0118
         Height          =   1875
         Left            =   -73080
         TabIndex        =   39
         Top             =   2280
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3307
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   1
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdUServicio"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoUniSer 
         Height          =   375
         Left            =   -74760
         Top             =   3360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin VB.CommandButton cmdDelUniSer 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   38
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtUniSer 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -72000
         MaxLength       =   30
         TabIndex        =   37
         Top             =   1560
         Width           =   3615
      End
      Begin VB.CommandButton cmdSaveUniSer 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   36
         Top             =   4680
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc adoUniPer 
         Height          =   375
         Left            =   -74760
         Top             =   3360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin VB.CommandButton cmdDelUniPer 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   27
         Top             =   4620
         Width           =   1335
      End
      Begin VB.TextBox txtUniPer 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -71880
         MaxLength       =   30
         TabIndex        =   26
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CommandButton cmdSaveUniPer 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   25
         Top             =   4620
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc adoSLaboral 
         Height          =   330
         Left            =   -74640
         Top             =   3120
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
      Begin VB.CommandButton cmdDelSitLab 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73080
         TabIndex        =   21
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtSitLab 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -72240
         MaxLength       =   30
         TabIndex        =   20
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdSaveSitLab 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71640
         TabIndex        =   19
         Top             =   4680
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc adoGrado 
         Height          =   330
         Left            =   -74760
         Top             =   3240
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
      Begin VB.CommandButton cmdDelGdo 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   15
         Top             =   4740
         Width           =   1335
      End
      Begin VB.TextBox txtGdo 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -72720
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdSaveGdo 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   13
         Top             =   4740
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc adoEstCiv 
         Height          =   375
         Left            =   -74640
         Top             =   2760
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin VB.CommandButton cmdDelEstCivil 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   9
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtEstCiv 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -72600
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdSaveEstCiv 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         TabIndex        =   7
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   4
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtcat 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   -72480
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "&Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   1
         Top             =   4680
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "fjVarios.frx":0130
         Height          =   1875
         Left            =   -73080
         TabIndex        =   2
         Top             =   2400
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdCatSoc"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adocatsocio 
         Height          =   375
         Left            =   -74880
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         UserName        =   "dpintos"
         Password        =   "cornolios"
         RecordSource    =   "CatSocio"
         Caption         =   "catsocio"
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
      Begin MSDataGridLib.DataGrid DGEstCiv 
         Bindings        =   "fjVarios.frx":014A
         Height          =   1875
         Left            =   -73080
         TabIndex        =   10
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdEstCiv"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGGdo 
         Bindings        =   "fjVarios.frx":0162
         Height          =   1875
         Left            =   -73080
         TabIndex        =   16
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdGrado"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGSitLab 
         Bindings        =   "fjVarios.frx":0179
         Height          =   1875
         Left            =   -73080
         TabIndex        =   22
         Top             =   2280
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3307
         _Version        =   393216
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdSitLab"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGUniPer 
         Bindings        =   "fjVarios.frx":0193
         Height          =   1875
         Left            =   -73200
         TabIndex        =   28
         Top             =   2400
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdUpertenece"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   -73680
         TabIndex        =   31
         Top             =   4380
         Width           =   4575
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   3000
            TabIndex        =   47
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   -73560
         TabIndex        =   32
         Top             =   4320
         Width           =   4935
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   3360
            TabIndex        =   46
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -73680
         TabIndex        =   33
         Top             =   4440
         Width           =   4455
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3000
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -73680
         TabIndex        =   34
         Top             =   4440
         Width           =   4335
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            TabIndex        =   44
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   -73680
         TabIndex        =   35
         Top             =   4560
         Width           =   4455
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   45
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Height          =   795
         Left            =   -73800
         TabIndex        =   42
         Top             =   4380
         Width           =   4935
         Begin VB.CommandButton cmdSalir 
            Appearance      =   0  'Flat
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   3360
            TabIndex        =   48
            Top             =   300
            Width           =   1335
         End
      End
      Begin MSAdodcLib.Adodc adoGrupo 
         Height          =   375
         Left            =   -74400
         Top             =   2640
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         RecordSource    =   "select * from tbl_grupos order by idRubro"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "fjVarios.frx":01AB
         Height          =   1875
         Left            =   -72960
         TabIndex        =   50
         Top             =   2280
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "idRubro"
            Caption         =   "Número"
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoFunc 
         Height          =   375
         Left            =   360
         Top             =   2820
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
         RecordSource    =   "TBL_Funcio"
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "fjVarios.frx":01C2
         Height          =   1875
         Left            =   1920
         TabIndex        =   67
         Top             =   1980
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "CodSeg"
            Caption         =   "Código"
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
            DataField       =   "NIVEL"
            Caption         =   "Nivel"
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
            DataField       =   "NOMBRE"
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
            DataField       =   "CLAVE"
            Caption         =   "Clave"
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
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915,024
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoTC 
         Height          =   375
         Left            =   -74640
         Top             =   2520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
         RecordSource    =   "TBL_TCambio"
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "fjVarios.frx":01D8
         Height          =   1875
         Left            =   -73080
         TabIndex        =   74
         Top             =   1680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3307
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "TC_Fecha"
            Caption         =   "TC_Fecha"
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
            DataField       =   "TC_Dolar"
            Caption         =   "TC_Dolar"
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
            DataField       =   "TC_Real"
            Caption         =   "TC_Real"
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
            DataField       =   "TC_PArg"
            Caption         =   "TC_PArg"
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
            DataField       =   "TC_UR"
            Caption         =   "TC_UR"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1065,26
            EndProperty
         EndProperty
      End
      Begin VB.Label Label15 
         Caption         =   "No.Orden"
         Height          =   315
         Index           =   1
         Left            =   -74520
         TabIndex        =   60
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Ult.Dia Presupuesto:"
         Height          =   375
         Index           =   0
         Left            =   -74520
         TabIndex        =   59
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Nuevo Grupo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74280
         TabIndex        =   54
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   53
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   41
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Nueva Unidad de Servicios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   1620
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74280
         TabIndex        =   30
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Nueva Unidad de Pertenencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   29
         Top             =   1740
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Nueva Situación Laboral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   23
         Top             =   1620
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   18
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Nuevo Grado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74280
         TabIndex        =   17
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nuevo Estado Civil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74400
         TabIndex        =   11
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Nueva Categoría"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "fjVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rstCategoria As Recordset
Attribute rstCategoria.VB_VarHelpID = -1
Dim WithEvents rstestciv As Recordset
Attribute rstestciv.VB_VarHelpID = -1
Dim WithEvents rstgrado As Recordset
Attribute rstgrado.VB_VarHelpID = -1
Dim WithEvents rstuserv As Recordset
Attribute rstuserv.VB_VarHelpID = -1
Dim WithEvents rstupert As Recordset
Attribute rstupert.VB_VarHelpID = -1
Dim WithEvents rstsitlab As Recordset
Attribute rstsitlab.VB_VarHelpID = -1
Dim WithEvents rstGrupos As Recordset
Attribute rstGrupos.VB_VarHelpID = -1
Dim WithEvents adoParam As Recordset
Attribute adoParam.VB_VarHelpID = -1






Private Sub cmdAgregaFunc_Click()
adoFunc.Recordset.MoveFirst
adoFunc.Recordset.AddNew

End Sub

Private Sub cmdBorraFunc_Click()
adoFunc.Recordset.Delete
adoFunc.Recordset.Update
End Sub

Private Sub cmdDelSitLab_Click()
'Situacion Laboral
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstsitlab = New Recordset
If rstsitlab.State = adStateOpen Then rstsitlab.Close
rstsitlab.Open "select * from slaboral", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstsitlab.Sort = "idsitlab"

If DGSitLab.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from slaboral where idsitlab = " & DGSitLab.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adoSLaboral.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdDelUniPer_Click()
'Unidad a la que Pertenece
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstupert = New Recordset
If rstupert.State = adStateOpen Then rstupert.Close
rstupert.Open "select * from unidadpert", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstupert.Sort = "idupertenece"

If DGUniPer.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from unidadpert where idupertenece = " & DGUniPer.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adoUniPer.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdDelUniSer_Click()
'Unidad Servicio
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstuserv = New Recordset
If rstuserv.State = adStateOpen Then rstuserv.Close
rstuserv.Open "select * from unidadserv", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstuserv.Sort = "iduservicio"

If DGUniSer.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from unidadserv where iduservicio = " & DGUniSer.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adoUniSer.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub



Private Sub cmdGrabar_Click()

End Sub

Private Sub cmdGuardarFunc_Click()
adoFunc.Recordset.Update
End Sub

Private Sub cmdsalir_Click(Index As Integer)
Unload Me
End Sub


Private Sub cmdGuardarGrupo_Click()
 'Grupos
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstGrupos = New Recordset
If rstGrupos.State = adStateOpen Then rstGrupos.Close
rstGrupos.Open "select * from tbl_Grupos", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstGrupos.Sort = "idrubro"
res = vbYes
    If rstGrupos.RecordCount <> 0 Then
    'Busca la grado para verificar que no exista
        rstGrupos.MoveFirst
        For i = 1 To rstGrupos.RecordCount
            If UCase(Trim(rstGrupos!Desc)) = UCase(Trim(txtGrupos.Text)) Then
                res = MsgBox("El Grado " & txtGrupos.Text & _
                    " ya existe en la Base de Datos" & Chr(13) & _
                    " ¿ Esta seguro que desea guardar este nuevo registro ?", _
                    vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstGrupos.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstGrupos.RecordCount <> 0 Then
            rstGrupos.MoveFirst
            codul = rstGrupos!idrubro
            For i = 1 To rstGrupos.RecordCount
                rstGrupos.MoveNext
                If i <> rstGrupos.RecordCount Then
                    If (rstGrupos!idrubro - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstGrupos!idrubro
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into tbl_Grupos values(" & _
            Val(cod + 1) & ", '" & Trim(txtGrupos.Text) & "')"
        comando.Execute
        MsgBox "El Grado " & txtGrupos.Text & _
            " se ha guardado satisfactoriamente", _
            vbInformation, "Círculo Policial"
        txtGrupos.Text = ""
        txtGrupos.SetFocus
        adoGrupo.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtGrupos.Text = ""
        txtGrupos.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub
Private Sub cmdSaveGdo_Click()
'Grado
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstgrado = New Recordset
If rstgrado.State = adStateOpen Then rstgrado.Close
rstgrado.Open "select * from grado", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstgrado.Sort = "idgrado"
res = vbYes
    If rstgrado.RecordCount <> 0 Then
    'Busca la grado para verificar que no exista
        rstgrado.MoveFirst
        For i = 1 To rstgrado.RecordCount
            If UCase(Trim(rstgrado!Desc)) = UCase(Trim(txtcat.Text)) Then
                res = MsgBox("El Grado " & txtGdo.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstgrado.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstgrado.RecordCount <> 0 Then
            rstgrado.MoveFirst
            codul = rstgrado!idgrado
            For i = 1 To rstgrado.RecordCount
                rstgrado.MoveNext
                If i <> rstgrado.RecordCount Then
                    If (rstgrado!idgrado - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstgrado!idgrado
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into grado values(" & Val(cod + 1) & ", '" & Trim(txtGdo.Text) & "')"
        comando.Execute
        MsgBox "El Grado " & txtGdo.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtGdo.Text = ""
        txtGdo.SetFocus
        adoGrado.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtGdo.Text = ""
        txtGdo.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub


Private Sub cmdBorrarGrupo_Click()
'Borrar grupos
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstGrupos = New Recordset
If rstGrupos.State = adStateOpen Then rstGrupos.Close
rstGrupos.Open "select * from tbl_grupos", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstGrupos.Sort = "idrubro"

If DataGrid2.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from tbl_grupos where idrubro = " & DataGrid2.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adoGrupo.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub
Private Sub cmdDelGdo_Click()
'Grado
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstgrado = New Recordset
If rstgrado.State = adStateOpen Then rstgrado.Close
rstgrado.Open "select * from grado", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstgrado.Sort = "idgrado"

If DGGdo.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from grado where idgrado = " & DGGdo.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adoGrado.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdGuardar_Click()
'CATEGORIA
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstCategoria = New Recordset
If rstCategoria.State = adStateOpen Then rstCategoria.Close
rstCategoria.Open "select * from catsocio", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstCategoria.Sort = "idcatsoc"
res = vbYes
    If rstCategoria.RecordCount <> 0 Then
    'Busca la categoria para verificar que no exista
        rstCategoria.MoveFirst
        For i = 1 To rstCategoria.RecordCount
            If UCase(Trim(rstCategoria!Desc)) = UCase(Trim(txtcat.Text)) Then
                res = MsgBox("La Categoría " & txtcat.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstCategoria.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstCategoria.RecordCount <> 0 Then
            rstCategoria.MoveFirst
            codul = rstCategoria!idcatsoc
            For i = 1 To rstCategoria.RecordCount
                rstCategoria.MoveNext
                If i <> rstCategoria.RecordCount Then
                    If (rstCategoria!idcatsoc - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstCategoria!idcatsoc
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into catsocio values(" & Val(cod + 1) & ", '" & Trim(txtcat.Text) & "')"
        comando.Execute
        MsgBox "La Categoría " & txtcat.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtcat.Text = ""
        txtcat.SetFocus
        adocatsocio.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtcat.Text = ""
        txtcat.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub
Private Sub cmdBorrar_Click()
'CATEGORIA
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstCategoria = New Recordset
If rstCategoria.State = adStateOpen Then rstCategoria.Close
rstCategoria.Open "select * from catsocio", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstCategoria.Sort = "idcatsoc"

If DataGrid1.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from catsocio where idcatsoc = " & DataGrid1.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adocatsocio.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub cmdSaveEstCiv_Click()
'Estado Civil
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstestciv = New Recordset
If rstestciv.State = adStateOpen Then rstestciv.Close
rstestciv.Open "select * from EstCivil", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstestciv.Sort = "idestciv"
res = vbYes
    If rstestciv.RecordCount <> 0 Then
    'Busca la categoria para verificar que no exista
        rstestciv.MoveFirst
        For i = 1 To rstestciv.RecordCount
            If UCase(Trim(rstestciv!Desc)) = UCase(Trim(txtEstCiv.Text)) Then
                res = MsgBox("El Estado Civil " & Me.txtEstCiv.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstestciv.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstestciv.RecordCount <> 0 Then
            rstestciv.MoveFirst
            codul = rstestciv!idestciv
            For i = 1 To rstestciv.RecordCount
                rstestciv.MoveNext
                If i <> rstestciv.RecordCount Then
                    If (rstestciv!idestciv - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstestciv!idestciv
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into EstCivil  values(" & Val(cod + 1) & ", '" & Trim(Me.txtEstCiv.Text) & "')"
        comando.Execute
        MsgBox "El Estado Civil" & Me.txtEstCiv.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtcat.Text = ""
        txtcat.SetFocus
        adoEstCiv.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtcat.Text = ""
        txtcat.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub
Private Sub cmdDelEstCivil_Click()
'Estado Civil
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstestciv = New Recordset
If rstestciv.State = adStateOpen Then rstestciv.Close
rstestciv.Open "select * from EstCivil", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstestciv.Sort = "idestciv"

If Me.DGEstCiv.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from EstCivil where idEstCiv = " & Me.DGEstCiv.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    Me.adoEstCiv.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdSaveSitLab_Click()
'situacion laboral
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstsitlab = New Recordset
If rstsitlab.State = adStateOpen Then rstsitlab.Close
rstsitlab.Open "select * from slaboral ", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstsitlab.Sort = "idsitlab"
res = vbYes
    If rstsitlab.RecordCount <> 0 Then
    'Busca la grado para verificar que no exista
        rstsitlab.MoveFirst
        For i = 1 To rstsitlab.RecordCount
            If UCase(Trim(rstsitlab!Desc)) = UCase(Trim(txtSitLab.Text)) Then
                res = MsgBox("La Situación Laboral " & txtSitLab.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstsitlab.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstsitlab.RecordCount <> 0 Then
            rstsitlab.MoveFirst
            codul = rstsitlab!idsitlab
            For i = 1 To rstsitlab.RecordCount
                rstsitlab.MoveNext
                If i <> rstsitlab.RecordCount Then
                    If (rstsitlab!idsitlab - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstsitlab!idsitlab
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into slaboral values(" & Val(cod + 1) & ", '" & Trim(txtSitLab.Text) & "')"
        comando.Execute
        MsgBox "La Situación Laboral " & txtSitLab.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtSitLab.Text = ""
        txtSitLab.SetFocus
        adoSLaboral.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtSitLab.Text = ""
        txtSitLab.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub

Private Sub cmdSaveUniPer_Click()
'Unidad pertenece
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstupert = New Recordset
If rstupert.State = adStateOpen Then rstupert.Close
rstupert.Open "select * from unidadpert", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstupert.Sort = "idupertenece"
res = vbYes
    If rstupert.RecordCount <> 0 Then
    'Busca la grado para verificar que no exista
        rstupert.MoveFirst
        For i = 1 To rstupert.RecordCount
            If UCase(Trim(rstupert!Desc)) = UCase(Trim(txtUniPer.Text)) Then
                res = MsgBox("La Unidad  " & txtUniPer.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstupert.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstupert.RecordCount <> 0 Then
            rstupert.MoveFirst
            codul = rstupert!idupertenece
            For i = 1 To rstupert.RecordCount
                rstupert.MoveNext
                If i <> rstupert.RecordCount Then
                    If (rstupert!idupertenece - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstupert!idupertenece
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into unidadpert values(" & Val(cod + 1) & ", '" & Trim(txtUniPer.Text) & "')"
        comando.Execute
        MsgBox "La Unidad " & txtUniPer.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtUniPer.Text = ""
        txtUniPer.SetFocus
        adoUniPer.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtUniPer.Text = ""
        txtUniPer.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub

Private Sub cmdSaveUniSer_Click()
'Unidad presta servicio
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstuserv = New Recordset
If rstuserv.State = adStateOpen Then rstuserv.Close
rstuserv.Open "select * from unidadserv", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstuserv.Sort = "iduservicio"
res = vbYes
    If rstuserv.RecordCount <> 0 Then
    'Busca la grado para verificar que no exista
        rstuserv.MoveFirst
        For i = 1 To rstuserv.RecordCount
            If UCase(Trim(rstuserv!Desc)) = UCase(Trim(txtUniSer.Text)) Then
                res = MsgBox("La Unidad de Servicio " & txtUniSer.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstuserv.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstuserv.RecordCount <> 0 Then
            rstuserv.MoveFirst
            codul = rstuserv!iduservicio
            For i = 1 To rstuserv.RecordCount
                rstuserv.MoveNext
                If i <> rstuserv.RecordCount Then
                    If (rstuserv!iduservicio - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstuserv!iduservicio
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into unidadserv values(" & Val(cod + 1) & ", '" & Trim(txtUniSer.Text) & "')"
        comando.Execute
        MsgBox "La Unidad de Servicio " & txtUniSer.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtUniSer.Text = ""
        txtUniSer.SetFocus
        adoUniSer.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtUniSer.Text = ""
        txtUniSer.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
adoTC.Recordset.Delete
adoTC.Recordset.Update

End Sub

Private Sub Command1_Click()
adoTC.Recordset.MoveFirst
adoTC.Recordset.AddNew

End Sub

Private Sub Command3_Click()
adoTC.Recordset.Update

End Sub

Private Sub DataGrid3_Change()
MsgBox "CAMBIO"
End Sub

Private Sub Form_Load()

    'toma parametros .....................................
    Dim sM As String
    Dim nMes, nAnio, ndia As Integer
    
    'NIVELES DE SEGURIDAD ................................
    If vpnNivelFuncionario = 6 Then
        'SSTab1.TabVisible(7) = False
        'SSTab1.TabVisible(8) = False
        SSTab1.Visible = True
    Else
        SSTab1.Visible = False
    End If
    Set adoParam = New ADODB.Recordset
    Set adoParam.ActiveConnection = adoconn
    adoParam.Open "SELECT * FROM TBL_Parametros", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    adoParam.MoveFirst
    sM = adoParam!PRM_Prspst
    nMes = CInt(Right(sM, 2))
    nAnio = CInt(Left(sM, 4))
    ndia = adoParam!PRM_prsphst
    mskUltimo.Mask = "##/##/####"
    mskUltimo.Text = CDate(ndia & "/" & nMes & "/" & nAnio)
    
    txtOrden.Text = adoParam!nroorden

End Sub
Private Sub cmdGrabarParam_Click()
Dim dFecha As Date
Dim sMom As String

dFecha = CDate(mskUltimo.Text)
adoParam!nroorden = txtOrden.Text
adoParam!PRM_prsphst = Day(dFecha)
If Month(dFecha) < 10 Then
    sMom = "0" & CStr(Month(dFecha))
Else
    sMom = CStr(Month(dFecha))
End If
adoParam!PRM_Prspst = Year(dFecha) & sMom
adoParam.Update
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    adoParam.Close
    Set adoParam = Nothing
    Set fjDatos = Nothing
End Sub

