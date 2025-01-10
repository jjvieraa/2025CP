VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmodcomer 
   Caption         =   "Modificar Comercios"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   7095
      Begin VB.TextBox txtComNom 
         DataField       =   "Razon"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   30
         Top             =   2280
         Width           =   2655
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
         Left            =   750
         TabIndex        =   17
         Top             =   6240
         Width           =   2175
      End
      Begin VB.CommandButton cmdmodifcom 
         Caption         =   "Actualizar"
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
         Left            =   3750
         TabIndex        =   16
         Top             =   6240
         Width           =   2295
      End
      Begin VB.CheckBox Convenio 
         Caption         =   "Convenio c/cuota mensual fija"
         DataField       =   "Convenio"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   4350
         TabIndex        =   15
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CheckBox Discrimina 
         Caption         =   "Discriminar Gastos"
         DataField       =   "Discrimina"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   2670
         TabIndex        =   14
         Top             =   5400
         Width           =   1575
      End
      Begin VB.ComboBox Desc 
         DataField       =   "Desc"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5190
         TabIndex        =   13
         Top             =   4680
         Width           =   735
      End
      Begin VB.ComboBox Cierre 
         DataField       =   "Cierre"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4710
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CheckBox Trab_Coop 
         Caption         =   "Trabaja c/socio Cooperador"
         DataField       =   "Trab_Coop"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   750
         TabIndex        =   11
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Afiliación"
         Height          =   855
         Left            =   750
         TabIndex        =   8
         Top             =   4200
         Width           =   3135
         Begin VB.OptionButton optcoop 
            Caption         =   "Cooperador"
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optadherido 
            Caption         =   "Adherido"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox Razon 
         DataField       =   "Razon"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   2295
         Width           =   2055
      End
      Begin VB.TextBox Direc 
         DataField       =   "Direc"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   2880
         Width           =   5175
      End
      Begin VB.ComboBox Rubro 
         DataField       =   "Rubro"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   750
         TabIndex        =   5
         Top             =   1455
         Width           =   1935
      End
      Begin VB.CommandButton cmdverdatos 
         Caption         =   "Ver datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker fech_ing 
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   -2147483635
         Format          =   24510465
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox Tel 
         DataField       =   "Tel"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   750
         TabIndex        =   18
         Top             =   3615
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
         Left            =   750
         TabIndex        =   19
         Top             =   735
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
         Left            =   2550
         TabIndex        =   20
         Top             =   3600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   750
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Cierre Día"
         Height          =   255
         Left            =   4800
         TabIndex        =   29
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Descuento (%)"
         Height          =   255
         Left            =   4830
         TabIndex        =   28
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   3870
         TabIndex        =   27
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   4350
         TabIndex        =   26
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   750
         TabIndex        =   25
         Top             =   2655
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Rubro"
         Height          =   255
         Left            =   750
         TabIndex        =   23
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   750
         TabIndex        =   22
         Top             =   3375
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº del Comercio"
         Height          =   255
         Left            =   750
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox auxcop 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "cooperador"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox aux 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "adherido"
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmmodcomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rstRubro As Recordset
Attribute rstRubro.VB_VarHelpID = -1
Dim adoComercios As New ADODB.Recordset
Dim adomodificar As New ADODB.Command


Private Sub cmdmodifcom_Click()
Dim codrubro As Integer
'On Error GoTo error

Set adoComercios.ActiveConnection = adoconn
If adoComercios.State = adStateOpen Then adoComercios.Close
adoComercios.Open "select * from TBL_Comercios where nro = " & Val(Nro.Text) & "", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

'Buscar código rubro
rstRubro.MoveFirst
For i = 1 To rstRubro.RecordCount
    If Trim(UCase(rstRubro!Desc)) = Trim(UCase(Rubro.Text)) Then
        codrubro = rstRubro!idrubro
        Exit For
    End If
rstRubro.MoveNext
Next i
If Rubro.Text = "" Then codrubro = 0
'Fin Buscar código rubro

If Me.optadherido.Value = True And Me.optcoop.Value = False Then
    Set adomodificar.ActiveConnection = adoconn
    adomodificar.CommandText = "update TBL_Comercios set NombCom = '" & txtComNom.Text & "', Rubro = " & codrubro & ", fech_ing = '" & Fech_ing.Value & "', razon = '" & Razon.Text & "', Direc = '" & Direc.Text & "', Tel = '" & Tel.Text & "', Ruc = '" & RUC.Text & "',Cierre = " & Val(Cierre.Text) & "',Tipo = '" & Trim(aux.Text) & "', Desc = " & Val(Desc.Text) & ", Trab_coop = '" & Trab_Coop & "', Discrimina = '" & Discrimina & "',Convenio = '" & Convenio & "' where nro = " & Val(Nro.Text) & " "
    adomodificar.Execute
    
    MsgBox "Registro Modificado", vbExclamation, "Circulo Policial"
Else
    adomodificar.CommandText = "update TBL_Comercios set nombcom = '" & Me.txtComNom.Text & "', Rubro = " & codrubro & ",Fech_ing = '" & Fech_ing.Value & "',Razon = '" & Razon.Text & "',Direc = '" & Direc.Text & "', Tel = '" & Tel.Text & "', Ruc = '" & RUC.Text & "',Cierre = " & Val(Cierre.Text) & ",Tipo = '" & Trim(auxcop.Text) & "',Desc = " & Val(Desc.Text) & ", Trab_coop = '" & Trab_Coop & "', Discrimina = '" & Discrimina & "',Convenio = '" & Convenio & "' where nro = " & Val(Nro.Text) & ""
    Set adomodificar.ActiveConnection = adoconn
    adomodificar.Execute
    MsgBox "Registro Modificado", vbExclamation, "Circulo Policial"
End If
txtComNom.Text = ""
Cierre.Text = ""
Razon.Text = ""
Rubro.Text = ""
RUC.Text = ""
Tel.Text = ""
Desc.Text = ""
Direc.Text = ""
Nro.Text = ""
'Me.Fech_Ing.Text = ""

'Me.optadherido.Value = False
'Me.optcoop.Value = False

Me.Trab_Coop.Value = vbunckecked
Me.Discrimina.Value = vbunckecked
Me.Convenio.Value = vbunckecked

'error:
'MsgBox ("ocurrio error")
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdverdatos_Click()
Dim codrubro As Integer
On Error GoTo error

Set adoComercios.ActiveConnection = adoconn
If adoComercios.State = adStateOpen Then adoComercios.Close
adoComercios.Open "select * from TBL_Comercios where nro = " & Me.Nro.Text & "", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
txtComNom.Text = adoComercios!nombcom
Cierre.Text = adoComercios!Cierre
Razon.Text = adoComercios!Razon
RUC.Text = adoComercios!RUC
Tel.Text = adoComercios!Tel
Desc.Text = adoComercios!Desc
Direc.Text = adoComercios!Direc
Nro.Text = adoComercios!Nro
Fech_ing.Value = adoComercios!Fech_ing
Rubro.Text = adoComercios!Rubro
'categoria
Set rstRubro = New Recordset
rstRubro.Open "select * from rubro", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstRubro.RecordCount <> 0 Then
        For i = 1 To rstRubro.RecordCount
            If adoComercios!Rubro = rstRubro!idrubro Then
                Rubro.Text = rstRubro!Desc
                Exit For
            Else
                rstRubro.MoveNext
            End If
        Next i
    Else
        msgtablas
    End If




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
MsgBox "No existe el comercio", vbCritical, "Circulo POlicial"
End Sub

Private Sub Command6_Click()
fjIngresos.Show
End Sub

Private Sub Form_Load()
 
 'Cargar Rubro
    Set rstRubro = New Recordset
    rstRubro.Open "select * from rubro", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstRubro.RecordCount <> 0 Then
        For i = 1 To rstRubro.RecordCount
            Rubro.AddItem (rstRubro!Desc)
            rstRubro.MoveNext
        Next i
    Else
        msgtablas
    End If
End Sub
