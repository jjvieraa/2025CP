VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form fjIngComercio 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Ingreso de COMERCIO"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCantRecibos 
      Alignment       =   1  'Right Justify
      DataField       =   "Razon"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3840
      TabIndex        =   39
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtValConv 
      Alignment       =   1  'Right Justify
      DataField       =   "Razon"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton optcoop 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cooperador"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   3060
      Width           =   1335
   End
   Begin VB.OptionButton optadherido 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adherido"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3060
      Width           =   1095
   End
   Begin VB.TextBox Nro 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   1095
   End
   Begin MSMask.MaskEdBox FechaIng 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   300
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox Desc 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Cierre 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2460
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   -120
      TabIndex        =   32
      Top             =   4620
      Width           =   5535
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
         Picture         =   "frmIngComercio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Picture         =   "frmIngComercio.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "frmIngComercio.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   375
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
         Picture         =   "frmIngComercio.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAlterar 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdgravar 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdincluir 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CheckBox Convenio 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Convenio c/cuota mensual fija"
      DataField       =   "Convenio"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3645
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox Discrimina 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Discriminar Gastos"
      DataField       =   "Discrimina"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox Trab_Coop 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Trabaja c/socio Cooperador"
      DataField       =   "Trab_Coop"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Razon 
      DataField       =   "Razon"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3255
      TabIndex        =   4
      Top             =   1035
      Width           =   2055
   End
   Begin VB.TextBox Direc 
      DataField       =   "Direc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   165
      TabIndex        =   5
      Top             =   1740
      Width           =   5175
   End
   Begin VB.ComboBox Grupo 
      DataField       =   "Grupos"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   1935
   End
   Begin VB.TextBox txtComNom 
      DataField       =   "Razon"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   1035
      Width           =   2775
   End
   Begin MSMask.MaskEdBox Tel 
      DataField       =   "Tel"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   165
      TabIndex        =   6
      Top             =   2475
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox RUC 
      DataField       =   "RUC"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1965
      TabIndex        =   7
      Top             =   2475
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cant.Recibos:"
      Height          =   255
      Left            =   3720
      TabIndex        =   38
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Valor Convenio:"
      Height          =   255
      Left            =   1440
      TabIndex        =   37
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cierre Día"
      Height          =   255
      Left            =   4125
      TabIndex        =   31
      Top             =   2235
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descuento (%)"
      Height          =   255
      Left            =   165
      TabIndex        =   30
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   3255
      TabIndex        =   29
      Top             =   795
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha de Ingreso"
      Height          =   255
      Left            =   3720
      TabIndex        =   28
      Top             =   75
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dirección"
      Height          =   255
      Left            =   165
      TabIndex        =   27
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R.U.C."
      Height          =   255
      Left            =   1965
      TabIndex        =   26
      Top             =   2235
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grupo"
      Height          =   255
      Left            =   1605
      TabIndex        =   25
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   165
      TabIndex        =   24
      Top             =   2235
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nº del Comercio"
      Height          =   255
      Left            =   165
      TabIndex        =   23
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nombre"
      Height          =   255
      Left            =   135
      TabIndex        =   22
      Top             =   795
      Width           =   2775
   End
End
Attribute VB_Name = "fjIngComercio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoC As New ADODB.Recordset

Dim cCOMR As New clsComercios
Dim funcao As String





'=======================================================
Private Sub Convenio_Click()
'=======================================================
If Convenio Then
    Label9.Visible = True
    txtValConv.Visible = True
Else
    Label9.Visible = False
    txtValConv.Visible = False
End If
End Sub

'=======================================================
Private Sub txtValConv_Validate(Cancel As Boolean)
'=======================================================
If Not IsNumeric(txtValConv.Text) Then
    Cancel = True
Else
    txtValConv.Text = Format(txtValConv.Text, "#,#0.00")
End If
End Sub

'=======================================================
Private Sub Desc_Validate(Cancel As Boolean)
'=======================================================
If Not IsNumeric(Desc.Text) Then
    Cancel = True
Else
    Desc.Text = Format(Desc.Text, "#,#0.00")
End If
End Sub
'=====================================================================
Private Sub txtValConv_GotFocus()
'=====================================================================
   
txtValConv.SelStart = 0
txtValConv.SelLength = Len(txtValConv.Text)
End Sub
'=====================================================================
Private Sub Desc_GotFocus()
'=====================================================================
   
Desc.SelStart = 0
Desc.SelLength = Len(Desc.Text)
End Sub
'=====================================================================
Private Sub Cierre_GotFocus()
'=====================================================================
   
Cierre.SelStart = 0
Cierre.SelLength = Len(Cierre.Text)
End Sub
'=======================================================
Private Sub Form_Load()
'=======================================================
    On Error GoTo mEr333
    
    'la base de datos--------------
    Set adoC.ActiveConnection = adoconn
    If adoC.State = adStateOpen Then adoC.Close
    adoC.Open "select * FROM tbl_Comercios ORDER BY codigo", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'carga grupos en el combobox GRUPO -------
    If Not cCOMR.CargaGruposEnCombo(Grupo) Then
        MsgBox "ERROR 342: No se pudo abrir la tabla de GRUPOS"
        Call cmdsalir_Click
    End If
    inicio
    Call cmdPrimeiro_Click
    Exit Sub
mEr333:
             MsgBox "ERROR 343: " & Err.Description & "  " & Err.Number
 Call cmdsalir_Click
End Sub



'=======================================================
Private Sub Form_Unload(Cancel As Integer)
'=======================================================

If adoC.State = adStateOpen Then
    adoC.Close
End If


Set adoC = Nothing
Set cCOMR = Nothing
Set fjIngComercio = Nothing
End Sub


'=======================================================
Private Sub cmdsalir_Click()
'=======================================================
Unload Me
End Sub



'=======================================================
Private Sub ActualizaCampos()
'=======================================================
    Dim codrubro As Integer
    Dim nI As Integer
    Dim sMome As String
    'On Error GoTo mEr334
    
    
   
    
    'AHERIDO
    If optadherido.Value = True Then
        sMome = "Adherido"
    Else
        sMome = "Cooperador"
    End If
    
    'ACTUALIZA
    
    adoC("Codigo") = Val(Nro.Text)
    adoC("NombCom") = txtComNom.Text
    adoC("grupo") = cCOMR.DevuelveNoRubro(Grupo.ListIndex)
    adoC("Fech_ing") = CDate(FechaIng.Text)
    adoC("razon") = Razon.Text
    adoC("direc") = Direc.Text
    adoC("tel") = Tel.Text
    adoC("ruc") = RUC.Text
    adoC("cierre") = CInt(Cierre.Text)
    If optadherido Then
        adoC("TIPO") = "A"
    Else
        adoC("tipo") = "C"
    End If
    adoC("desc") = CSng(Desc.Text)
    adoC("trab_coop") = Trab_Coop.Value
    adoC("discrimina") = Discrimina.Value
    adoC("CONVENIO") = Convenio.Value
    adoC("ConvCuota") = CSng(txtValConv.Text)
    adoC("CantRecib") = CInt(txtCantRecibos.Text)        'CANTIDAD RECIBOS A IMPRIMIR
    adoC("FUNC") = vpnFuncionario
    adoC("DIA") = Date
    adoC("HORA") = Time
    
    Exit Sub


mEr334:
    MsgBox "ERROR 2344: " & Err.Description & "  No: " & Err.Number
End Sub



'=======================================================
Private Sub ActualizaFormulario()
'=======================================================
Nro.Text = adoC("Codigo")
Grupo.ListIndex = cCOMR.DevuelveNoDelCombo(adoC("grupo"))
EnmCmpToObjF FechaIng, adoC("fech_ing")
txtComNom.Text = "" & adoC("nombcom")
Razon.Text = "" & adoC("razon")
Direc.Text = "" & adoC("direc")
Tel.Text = "" & adoC("tel")
RUC.Text = "" & adoC("ruc")
Cierre.Text = EnmCmpToObjN(adoC("cierre"))
If adoC("tipo") = "A" Then
    optadherido.Value = True
Else
    optcoop.Value = True
End If
Desc.Text = Format(adoC("desc"), "#,#0.00")
txtCantRecibos.Text = adoC("CantRecib")
txtValConv.Text = Format(adoC("ConvCuota"), "#,#0.00")
Trab_Coop.Value = IIf(adoC("trab_coop"), vbChecked, vbUnchecked)
Discrimina.Value = IIf(adoC("discrimina"), vbChecked, vbUnchecked)
Convenio.Value = IIf(adoC("convenio"), vbChecked, vbUnchecked)
End Sub



'=======================================================
Private Sub Nro_KeyDown(KeyCode As Integer, Shift As Integer)
'=======================================================
    If KeyCode = 113 Then   'F2
        vpMuestraTabla = kMuestraComercios
        fjMuestraTabla.Show
    End If
End Sub












'000000000000000000000000000000000000000000000000
' B A RR A       DE       T A R E A S
'0000000000000000000000000000000000000000000000000

Private Sub cmdAlterar_Click()
    funcao = "ALT"
    botoes
    ActualizaFormulario
End Sub

Private Sub cmdCancelar_Click()
    If adoC.RecordCount > 0 Then
        adoC.MoveFirst
        ActualizaFormulario
    Else
        LimpiaBoxes
    End If
    inicio

End Sub

Private Sub cmdExcluir_Click()
    funcao = "EXC"
    If MsgBox("Confirma ?", vbYesNo, "Confirmando !") = vbYes Then
            adoC.Delete
            If adoC.RecordCount > 0 Then
            ' Mostra o registro anterior pois esse nao existe mais
                cmdAnterior_Click
            Else
                LimpiaBoxes
            End If
    End If
    inicio

End Sub
'================================
Private Function ValidaCampos() As Boolean
'================================
If Nro.Text = "" Or Not IsNumeric(Nro.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    Nro.SetFocus
    ValidaCampos = False
    Exit Function
End If
If txtComNom.Text = "" Or IsNull(txtComNom.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    txtComNom.SetFocus
    ValidaCampos = False
    Exit Function
End If
'If Razon.Text = "" Or IsNull(Razon.Text) Then
'    MsgBox "Campo incompleto", 16, "Aviso"
'    Razon.SetFocus
'    ValidaCampos = False
'    Exit Function
'End If
If Not IsDate(FechaIng.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    FechaIng.SetFocus
    ValidaCampos = False
    Exit Function
End If
If Direc.Text = "" Or IsNull(Direc.Text) Then
    MsgBox "Campo incompleto", 16, "Aviso"
    Direc.SetFocus
    ValidaCampos = False
    Exit Function
End If
If txtCantRecibos.Text = "" Or _
    IsNull(txtCantRecibos.Text) Or _
    Not IsNumeric(txtCantRecibos.Text) Then
    MsgBox "Campo incorrecto", 16, "Aviso"
    txtCantRecibos.SetFocus
    ValidaCampos = False
    Exit Function
End If
If Desc.Text = "" Or IsNull(Desc.Text) Then
    Desc.Text = "0"
End If
If txtValConv.Text = "" Or IsNull(txtValConv.Text) Then
    txtValConv.Text = "0"
End If
txtComNom.Text = UCase(txtComNom.Text)
Razon.Text = UCase(Razon.Text)
Direc.Text = UCase(Direc.Text)
ValidaCampos = True
End Function



'================================
Private Sub cmdgravar_Click()
'================================
    If Not ValidaCampos Then
        Exit Sub
    End If
    ' Critica os dados do cliente a nivel de Banco de Dados
    If funcao = "INC" Then
        If adoC.RecordCount > 0 Then
            adoC.MoveFirst
            adoC.Find "Codigo = " & Nro.Text
            If Not adoC.EOF Then
                ' ya existe  codigo
                MsgBox "Ya existe el Código", vbCritical, "Aviso"
                Nro.Text = ""
                Nro.SetFocus
                Exit Sub
            End If      'not adoC.eof
        End If          'recordCount
        adoC.AddNew
        ActualizaCampos
        adoC.Update
        inicio
    End If              ' funcao INC
        
    If funcao = "ALT" Then
        ActualizaCampos
        adoC.Update
        'MsgBox rsMsg("msg3"), vbOKOnly, "Aviso"
        
        'para localizar o registro
        inicio
        
    End If
End Sub

'================================
Private Sub cmdincluir_Click()
'================================
    funcao = "INC"
    botoes
    LimpiaBoxes
    'toma el ultimo No para aumentarlo
    adoC.MoveLast
    Nro.Text = adoC("codigo") + 1
    FechaIng.Text = Date
    Nro.SetFocus
End Sub

'================================
Private Sub cmdSair_Click()
'================================
   Unload Me

End Sub

'================================
Private Sub cmdPrimeiro_Click()
'================================
    If adoC.RecordCount > 0 Then
        adoC.MoveFirst
        If adoC.BOF = True Then
            adoC.MoveFirst
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
    If adoC.RecordCount > 0 Then
        adoC.MoveNext
        If adoC.EOF = True Then
            adoC.MovePrevious
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
    If adoC.RecordCount > 0 Then
    adoC.MovePrevious
    If adoC.BOF = True Then
        adoC.MoveNext
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
    If adoC.RecordCount > 0 Then
        adoC.MoveLast
        If adoC.EOF = True Then
            adoC.MoveLast
        End If
        ActualizaFormulario
    Else
        MsgBox "Sin Registros", 16, "Aviso"
        Exit Sub
    End If
End Sub
'================================
Private Sub botoes()
'================================
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


'================================
Sub inicio()
'================================
    adoC.Requery
    If adoC.RecordCount > 0 Then
        If funcao <> "ALT" Then
            adoC.MoveFirst
            ActualizaFormulario
        End If
    Else
        LimpiaBoxes
    End If
    
    cmdincluir.Enabled = True
    cmdSair.Enabled = True
    cmdgravar.Enabled = False
    cmdCancelar.Enabled = False
    
    
    If adoC.RecordCount = 0 Then
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


'================================
Private Sub LimpiaBoxes()
'================================
  'CERA LOS BOXES
    Nro.Text = ""
    Grupo.ListIndex = 0
    EnmInicioF FechaIng
    txtComNom.Text = ""
    Cierre.Text = "0"
    Razon.Text = ""
    RUC.Text = ""
    Tel.Text = ""
    Desc.Text = "0,00"
    Direc.Text = ""
    Nro.Text = ""
    optadherido.Value = True
    optcoop.Value = False
    Trab_Coop.Value = vbUnchecked
    Discrimina.Value = vbUnchecked
    Convenio.Value = vbUnchecked
    txtValConv.Text = "0,00"
End Sub


'================================
Private Function EnmObjToCmpF(mPrm As MaskEdBox) As Date
'================================
'Enmascara Objeto to Campo
'campo = box.txt
    If mPrm.Text = "__/__/____" Then
         EnmObjToCmpF = CDate("0")
    Else
        EnmObjToCmpF = mPrm.Text
    End If
End Function
'================================
Private Function EnmObjToCmpN(tPrm As TextBox) As Currency
'================================
    
    If tPrm.Text = "" Then
         EnmObjToCmpN = 0
    Else
        EnmObjToCmpN = tPrm.Text
    End If
End Function
'================================
Private Sub EnmInicioF(mPrm As MaskEdBox)
'================================
'ojo a la sintaxis   enminiciof nombre_del_meb
    mPrm.Mask = ""
    mPrm.Text = ""
    mPrm.Mask = "##/##/####"
End Sub
'================================
Private Function EnmCmpToObjN(sPrm As Variant) As Currency
'================================
'box.text = campo
   If sPrm = "" Or IsNull(sPrm) Then
         EnmCmpToObjN = 0
    Else
        EnmCmpToObjN = sPrm
    End If

End Function




'================================
Private Sub EnmCmpToObjF(mPrm As MaskEdBox, dPrm As Variant)
'================================
'objeto.text = campo

If IsNull(dPrm) Or _
    dPrm = "__/__/____" Or _
    dPrm = CDate("0") Then
        mPrm.Mask = ""
        mPrm.Text = ""
        mPrm.Mask = "##/##/####"
    Else
        mPrm.Text = Format(dPrm, "short date")
    End If
End Sub


