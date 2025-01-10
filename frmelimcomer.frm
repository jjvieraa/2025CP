VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmelimcomer 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox aux 
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Text            =   "adherido"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton verdatos 
      Caption         =   "Ver Datos"
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox rubro 
      Height          =   315
      Left            =   3240
      TabIndex        =   24
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox cierre 
      Height          =   315
      Left            =   7200
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox desc 
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox Direc 
      DataField       =   "Direc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3270
      TabIndex        =   8
      Top             =   3000
      Width           =   5175
   End
   Begin VB.TextBox Razon 
      DataField       =   "Razon"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3240
      TabIndex        =   7
      Top             =   2295
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Afiliación"
      Height          =   855
      Left            =   3270
      TabIndex        =   5
      Top             =   4320
      Width           =   3135
      Begin VB.OptionButton optadherido 
         Caption         =   "Adherido"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optcoop 
         Caption         =   "Cooperador"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CheckBox Trab_Coop 
      Caption         =   "Trabaja c/socio Cooperador"
      DataField       =   "Trab_Coop"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3270
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox Discrimina 
      Caption         =   "Discriminar Gastos"
      DataField       =   "Discrimina"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5190
      TabIndex        =   3
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox Convenio 
      Caption         =   "Convenio c/cuota mensual fija"
      DataField       =   "Convenio"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6870
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdelimcom 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6270
      TabIndex        =   1
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   6360
      Width           =   2175
   End
   Begin MSMask.MaskEdBox Tel 
      DataField       =   "Tel"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3270
      TabIndex        =   9
      Top             =   3735
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Fech_Ing 
      DataField       =   "Fech_Ing"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   5550
      TabIndex        =   10
      Top             =   1695
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Nro 
      DataField       =   "Nro"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3270
      TabIndex        =   11
      Top             =   855
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
      Left            =   5070
      TabIndex        =   22
      Top             =   3735
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Nº del Comercio"
      Height          =   255
      Left            =   3270
      TabIndex        =   20
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   3270
      TabIndex        =   19
      Top             =   3495
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Rubro"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "R.U.C."
      Height          =   255
      Left            =   5070
      TabIndex        =   17
      Top             =   3495
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   3270
      TabIndex        =   16
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de Ingreso"
      Height          =   255
      Left            =   5550
      TabIndex        =   15
      Top             =   1455
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   2055
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Descuento (%)"
      Height          =   255
      Left            =   7350
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Cierre Día"
      Height          =   255
      Left            =   7230
      TabIndex        =   12
      Top             =   3495
      Width           =   855
   End
End
Attribute VB_Name = "frmelimcomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoComercios As New ADODB.Recordset
Dim adoeliminar As New ADODB.Command

Private Sub cmdelimcom_Click()

Set adoComercios.ActiveConnection = adoconn
If adoComercios.State = adStateOpen Then adoComercios.Close
adoComercios.Open "select * from TBL_Comercios  where nro = " & Me.Nro.Text & "", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

adoeliminar.CommandText = "delete from TBL_Comercios where nro = " & Me.Nro.Text & ""
Set adoeliminar.ActiveConnection = adoconn
adoeliminar.Execute
MsgBox "Registro Eliminado", vbExclamation, "Circulo Policial"

Cierre.Text = ""
Razon.Text = ""
Rubro.Text = ""
RUC.Text = ""
Tel.Text = ""
Desc.Text = ""
Direc.Text = ""
Nro.Text = ""
Me.Fech_ing.Text = "__/__/____"


'Me.optadherido.Value = False
'Me.optcoop.Value = False

Trab_Coop.Value = vbunckecked
Discrimina.Value = vbunckecked
Convenio.Value = vbunckecked
End Sub


Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub verdatos_Click()
On Error GoTo error

Set adoComercios.ActiveConnection = adoconn
If adoComercios.State = adStateOpen Then adoComercios.Close

adoComercios.Open "select * from TBL_Comercios where nro = " & Me.Nro.Text & "", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

Cierre.Text = adoComercios!Cierre
Razon.Text = adoComercios!Razon
Rubro.Text = adoComercios!Rubro
RUC.Text = adoComercios!RUC
Tel.Text = adoComercios!Tel
Desc.Text = adoComercios!Desc
Direc.Text = adoComercios!Direc
Nro.Text = adoComercios!Nro
Fech_ing.Text = adoComercios!Fech_ing

If UCase(Trim(adoComercios!tipo)) = UCase(Trim(aux.Text)) Then
   Me.optadherido.Value = True
   Me.optcoop.Value = False
Else
   Me.optadherido.Value = False
   Me.optcoop.Value = True
End If
If adoComercios!Trab_Coop = 0 Then
   Me.Trab_Coop.Value = vbUnchecked
Else
Me.Trab_Coop.Value = vbChecked
End If
If adoComercios!Discrimina = 0 Then
Me.Discrimina.Value = vbUnchecked
Else
Me.Discrimina.Value = vbChecked
End If
If adoComercios!Convenio = 0 Then
Me.Convenio.Value = vbUnchecked
Else
Me.Convenio.Value = vbChecked
End If
Exit Sub

error:
MsgBox "No existe el comercio", vbCritical, "Circulo Policial"
End Sub
