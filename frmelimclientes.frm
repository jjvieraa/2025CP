VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmelimclientes 
   Caption         =   "Eliminar Socio"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Datos Laborales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   3975
      Left            =   6000
      TabIndex        =   26
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox ocupacion 
         DataField       =   "Ocupacion"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtgrado 
         DataField       =   "Grado"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1935
      End
      Begin VB.TextBox txtuservicios 
         DataField       =   "U_Servicio"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox txtupertenece 
         DataField       =   "U_Pertenece"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2880
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2925
         Width           =   2415
      End
      Begin VB.TextBox txtstlab 
         DataField       =   "SitLab"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2490
         Width           =   2415
      End
      Begin VB.TextBox txtcategoria 
         DataField       =   "Categoria"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2415
      End
      Begin MSMask.MaskEdBox ingresos 
         DataField       =   "Ingresos"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1920
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   705
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
      Begin MSMask.MaskEdBox Limite 
         DataField       =   "Limite"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1920
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1170
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
         Caption         =   "Limite de Credito"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Ocupación"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Ingresos"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Grado"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Unidad donde presta Servicios"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   3390
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Unidad a la que Pertenece"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   2955
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Categoría"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   2085
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Situación Labolal"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   2520
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00404000&
      Height          =   3975
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox Apellido 
         DataField       =   "Apellido"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   480
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox localidad 
         DataField       =   "Localidad"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   480
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox direccion 
         DataField       =   "Direccion"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2520
         Width           =   4815
      End
      Begin VB.TextBox nombre 
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtest_civ 
         DataField       =   "Est_Civil"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1695
      End
      Begin MSMask.MaskEdBox ci 
         DataField       =   "CI"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSMask.MaskEdBox Fech_Nac 
         DataField       =   "Fech_Nac"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   480
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label23 
         Caption         =   "Apellidos"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Estado Civil"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Cedula de Identidad"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Nacimiento"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nombres"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdmostrar 
      Caption         =   "Ver Datos"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fech_Ing 
      DataField       =   "Fech_Ing"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   9360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox NroSoc 
      DataField       =   "NroSoc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
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
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   25
      Top             =   6480
      Width           =   4935
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdsoselimina 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de Ingreso"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Nº de Cobro"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nº del Socio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label aclara 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3315
      TabIndex        =   4
      Top             =   7740
      Width           =   1770
   End
End
Attribute VB_Name = "frmelimclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoclientes As New ADODB.Recordset
Dim adoeliminar As New ADODB.Command
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

Private Sub cmdmostrar_Click()
Dim i As Integer
On Error GoTo error

Set adoclientes.ActiveConnection = adoconn
If adoclientes.State = adStateOpen Then adoclientes.Close

adoclientes.Open "select * from socios where nrosoc = " & Val(Me.NroSoc.Text) & ""
If adoclientes.RecordCount = 0 Then
    MsgBox "No existe el Socio", vbCritical, "Registro Inexistente"
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
    fech_ing.Text = adoclientes!fech_ing
    Fech_nac.Text = adoclientes!Fech_nac
    ci.Text = adoclientes!ci

'////**** búsqueda de códigos***/////

'categoria
Set rstcatsoc = New Recordset
rstcatsoc.Open "select * from catsocio", adoconn
    If rstcatsoc.RecordCount <> 0 Then
        For i = 1 To rstcatsoc.RecordCount
            If adoclientes!codcatsoc = rstcatsoc!idcatsoc Then
                txtcategoria.Text = rstcatsoc!Desc
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
rstgrado.Open "select * from grado", adoconn
    If rstgrado.RecordCount <> 0 Then
        For i = 1 To rstgrado.RecordCount
            If adoclientes!codgrado = rstgrado!idgrado Then
                txtgrado.Text = rstgrado!Desc
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
rstestciv.Open "select * from estcivil", adoconn
    If rstestciv.RecordCount <> 0 Then
        For i = 1 To rstestciv.RecordCount
            If adoclientes!codestciv = rstestciv!idestciv Then
                txtest_civ.Text = rstestciv!Desc
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
rstsitlab.Open "select * from SLaboral", adoconn
    If rstsitlab.RecordCount <> 0 Then
        For i = 1 To rstsitlab.RecordCount
            If adoclientes!codsitlab = rstsitlab!idsitlab Then
                txtstlab.Text = rstsitlab!Desc
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
rstupert.Open "select * from unidadpert", adoconn
    If rstupert.RecordCount <> 0 Then
        For i = 1 To rstupert.RecordCount
            If adoclientes!codunidper = rstupert!idupertenece Then
                txtupertenece.Text = rstupert!Desc
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
rstprestserv.Open "select * from unidadserv", adoconn
    If rstprestserv.RecordCount <> 0 Then
        For i = 1 To rstprestserv.RecordCount
            If adoclientes!codpresserv = rstprestserv!iduservicio Then
                txtuservicios.Text = rstprestserv!Desc
                Exit For
            Else
                rstprestserv.MoveNext
            End If
        Next i
    Else
        msgtablas
    End If
'Me.Cred_Auto = adoclientes!Cred_Auto
Me.Limite.Text = adoclientes!Limite
End If
Exit Sub

error:
    MsgBox Err.Description
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdsoselimina_Click()
Dim res As Integer
Set adoclientes.ActiveConnection = adoconn
If adoclientes.State = adStateOpen Then adoclientes.Close
adoclientes.Open "select * from socios  where nrosoc = " & Val(Me.NroSoc.Text) & ""

res = MsgBox("Está Seguro que desea eliminar este registro", vbYesNoCancel + vbQuestion, "Registros")
If res = vbYes Then
    adoeliminar.CommandText = "delete from socios where nrosoc = " & Val(Me.NroSoc.Text) & ""
    Set adoeliminar.ActiveConnection = adoconn
    adoeliminar.Execute
    MsgBox "Registro Eliminado", vbInformation, "Resgistros"
    txtupertenece.Text = ""
    txtuservicios.Text = ""
    Tel.Text = ""
    ocupacion.Text = ""
    NroSoc.Text = ""
    NroCob.Text = ""
    nombre.Text = ""
    Apellido.Text = ""
    localidad.Text = ""
    ingresos.Text = ""
    txtgrado.Text = ""
    txtest_civ.Text = ""
    direccion.Text = ""
    txtcategoria.Text = ""
    txtstlab.Text = ""
    Limite.Text = ""
    fech_ing.Text = "__/__/____"
    Fech_nac.Text = "__/__/____"
    ci.Text = "_.___.___-_"
    Exit Sub
Else
     If res = vbNo Then
        MsgBox "El registro no se ha eliminado", vbExclamation, "Registro"
     Else
        If res = vbCancel Then
            txtupertenece.Text = ""
            txtuservicios.Text = ""
            Tel.Text = ""
            ocupacion.Text = ""
            NroSoc.Text = ""
            NroCob.Text = ""
            nombre.Text = ""
            Apellido.Text = ""
            localidad.Text = ""
            ingresos.Text = ""
            txtgrado.Text = ""
            txtest_civ.Text = ""
            direccion.Text = ""
            txtcategoria.Text = ""
            txtstlab.Text = ""
            Limite.Text = ""
            fech_ing.Text = "__/__/____"
            Fech_nac.Text = "__/__/____"
            ci.Text = "_.___.___-_"
        End If
    End If
End If
End Sub

Private Sub Form_Load()


txtupertenece.Enabled = False
txtuservicios.Enabled = False
Tel.Enabled = False
ocupacion.Enabled = False
NroCob.Enabled = False
nombre.Enabled = False
Apellido.Enabled = False
localidad.Enabled = False
ingresos.Enabled = False
txtgrado.Enabled = False
txtest_civ.Enabled = False
direccion.Enabled = False
txtcategoria.Enabled = False
txtstlab.Enabled = False
fech_ing.Enabled = False
Fech_nac.Enabled = False
ci.Enabled = False
Limite.Enabled = False

txtupertenece.Text = ""
txtuservicios.Text = ""
Tel.Text = ""
ocupacion.Text = ""
NroSoc.Text = ""
NroCob.Text = ""
nombre.Text = ""
Apellido.Text = ""
localidad.Text = ""
ingresos.Text = ""
txtgrado.Text = ""
txtest_civ.Text = ""
direccion.Text = ""
txtcategoria.Text = ""
txtstlab.Text = ""
fech_ing.Text = "__/__/____"
Fech_nac.Text = "__/__/____"
ci.Text = "_.___.___-_"
Limite.Text = ""

End Sub




