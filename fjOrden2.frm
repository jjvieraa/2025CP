VERSION 5.00
Begin VB.Form fjOrden2 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes"
   ClientHeight    =   6225
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5775
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_FVto"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   51
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_FVto"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   44
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ORD_Cerro"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   38
      Top             =   3330
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   5535
      Begin VB.CommandButton cmdIncluir 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAlterar 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   32
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
         Picture         =   "fjOrden2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
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
         Picture         =   "fjOrden2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "fjOrden2.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "fjOrden2.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H0000FF00&
         Caption         =   "&Anular"
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
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ORD_MEPagos"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   26
      Top             =   3657
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Mon"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   23
      Top             =   1368
      Width           =   255
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ORD_EntCta"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   21
      Top             =   3003
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ORD_Recarg"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   19
      Top             =   2676
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_CtasPagas"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   17
      Top             =   2349
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Plan"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   15
      Top             =   2022
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_FVto"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_FEmis"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   11
      Top             =   3984
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ORD_Cuota"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   9
      Top             =   1695
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Depend"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1041
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_NroOrden"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   714
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_NroCom"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   387
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_NroSoc"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Func:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   50
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3960
      TabIndex        =   49
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3960
      TabIndex        =   48
      Top             =   3033
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3960
      TabIndex        =   47
      Top             =   2706
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3960
      TabIndex        =   46
      Top             =   2379
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3960
      TabIndex        =   45
      Top             =   2052
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tipo:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   43
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   " "
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   41
      Top             =   1395
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   40
      Top             =   417
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   " "
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   39
      Top             =   90
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ME Pagos"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   25
      Top             =   3735
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ME Cuota:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Moneda:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   1305
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ent. a Cuenta:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3015
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Recargos:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cuotas Pagas:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2385
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Plan: "
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2055
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fecha Vencimiento:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   4380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fecha Emisión:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   4065
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Valor Cuota:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nro.Dependiente:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nro.Orden:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nro.Comercio:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nro.Socio:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "fjOrden2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'utilizo este form para mostrar una orden
'y para anular una orden
Option Explicit

Dim adoC As New ADODB.Recordset
Dim adoD As New ADODB.Recordset
Dim mQue As String
Dim cSOC As New clsSocios
Dim cCom As New clsComercios
Dim cPag As New clsPagos
Dim cRecib As New clsRecibos
Dim cTC As New clsTCambio



Private Sub cmdAnular_Click()
On Error GoTo VVV4
'se va a anular una orden
'0) Pide la causa
'fjPideDato.Text1.MultiLine = True
'fjPideDato.Text1.ScrollBars = 2
fjPideDato.Text1.MaxLength = 50
fjPideDato.Caption = "Motivo de la anulación:"
fjPideDato.Show vbModal

'1) Guarda la acción
msRegistraUnaAccion 4, txtFields(2).Text, vpsPideDato, vpnFuncionario, Date, Time

'2) Cancela la orden
Dim sValor As Single
Dim sMEenP As Single
Dim sTC As Single
sValor = CSng(txtFields(5).Text) * (CInt(txtFields(6).Text) - _
    CInt(txtFields(7).Text)) + CSng(txtFields(8).Text) - CSng(txtFields(9).Text)
adoC("ord_EntCta") = CSng(txtFields(9).Text) + sValor
adoC("ORD_CERRO") = Date
adoC("ord_Tipo") = 4
'SI ES EN MONEDA EXTRANJERA
If Not txtFields(4).Text = "P" Then
    sTC = cTC.mfDevuelveCambio(txtFields(4).Text, Date)
    sMEenP = sValor * sTC
Else
    sMEenP = 0
End If
adoC("ord_MEPagos") = sMEenP
adoC.Update
adoC.Close

cRecib.mfTomaNroRecibo
'3) Registra el pago
If cPag.mfAbrePagos Then
    If cPag.mfGuardaUnPago(CLng(txtFields(0).Text), CLng(txtFields(2).Text), _
        sValor, Format(Date, "short date"), vplNroRecibo, CLng(txtFields(1).Text), txtFields(4).Text, _
        CSng(txtFields(10).Text), 4, "Anulac", Format(Time, "short time"), CStr(vpnFuncionario)) Then
    End If
End If

'actualiza el numero de recibo
vplNroRecibo = vplNroRecibo + 1
cRecib.mfGuardaNroRecibo

Unload Me
Exit Sub
VVV4:
MsgBox "ERROR vvv4: Anulando Orden " & Err.Description & "  " & Err.Number
End Sub

'=======================================================
Private Sub Form_Load()
'=======================================================
    On Error GoTo mEr333
    
    'la base de datos--------------
    Set adoC.ActiveConnection = adoconn
    
    If Not cSOC.mfAbreTablaSociosOrdenSocio Then
        Unload Me
    End If
    If Not cCom.mfAbreTablaComercios Then
        Unload Me
    End If
    
   If adoC.State = adStateOpen Then adoC.Close
    
    If vpFormMovim = kFormRecorre Then
        If adoC.State = adStateOpen Then adoC.Close
        adoC.Open "select * FROM tbl_Ordenes ORDER BY ORD_nroORDEN;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

        Call inicio
        Call BloqueaTexto
        Call cmdPrimero_Click
        
        
        'SEGURIDAD...........................
        If Not vpnFuncionario = kYO Then
            cmdIncluir.Enabled = False
            cmdExcluir.Enabled = False
            cmdAlterar.Enabled = False
        End If
    ElseIf vpFormMovim = kFormAnula Then
        'Lo muestra
        adoC.Open "select * FROM tbl_Ordenes WHERE ORD_NroOrden =" & CLng(vpsPideDato), adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        If adoC.RecordCount = 1 Then
            Call inicio
            Call BloqueaTexto
            'Coloca el boton ANULAR
            'Y DESHABILITA TODOS LOS DEMAS
            cmdIncluir.Visible = False
            cmdAnular.Visible = True
            cmdExcluir.Enabled = False
            cmdAlterar.Enabled = False
        Else
            MsgBox "No existe la orden " & vpsPideDato
        End If
    ElseIf vpFormMovim = kFormMira Then
        'Lo muestra
        adoC.Open "select * FROM tbl_Ordenes WHERE ORD_NroOrden =" & CLng(vpsPideDato), adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        If adoC.RecordCount = 1 Then
            Call inicio
            Call BloqueaTexto
            'Coloca el boton ANULAR
            'Y DESHABILITA TODOS LOS DEMAS
            Frame3.Visible = False
        Else
            MsgBox "No existe la orden " & vpsPideDato
        End If
    End If
 Exit Sub
mEr333:
             MsgBox "ERROR a343: " & Err.Description & "  " & Err.Number
 Call cmdsalir_Click
End Sub



'=======================================================
Private Sub Form_Unload(Cancel As Integer)
'=======================================================

If adoC.State = adStateOpen Then
    adoC.Close
End If


Set adoC = Nothing
Set cSOC = Nothing
Set cCom = Nothing
Set cPag = Nothing
Set cRecib = Nothing
Set fjOrden2 = Nothing

End Sub



'================================
Private Sub cmdGrabar_Click()
'================================
    If Not ValidaCampos Then
        Exit Sub
    End If
    If mQue = "INC" Then
        adoC.AddNew
        ActualizaCampos
        adoC.Update
        inicio
    End If
    If mQue = "ALT" Then
        ActualizaCampos
        adoC.Update
        'MsgBox rsMsg("msg3"), vbOKOnly, "Aviso"
        
        'para localizar o registro
        inicio
    End If
    BloqueaTexto
End Sub


'================================
Private Sub cmdincluir_Click()
'================================
    'incluir un socio
    mQue = "INC"
    Botones
    LimpiaBoxes
    DesbloqueaTexto
    'NroSoc.Enabled = True
    'toma el ultimo No para aumentarlo
    adoC.MoveLast
End Sub


'================================
Private Sub cmdsalir_Click()
'================================
   Unload Me

End Sub


Private Sub cmdAlterar_Click()
    mQue = "ALT"
    DesbloqueaTexto
    Botones
    ActualizaFormulario
End Sub


Private Sub cmdCancelar_Click()
    mQue = ""
    BloqueaTexto
    If adoC.RecordCount > 0 Then
        adoC.MoveFirst
        ActualizaFormulario
    Else
        LimpiaBoxes
    End If
    inicio

End Sub


Private Sub cmdExcluir_Click()
    mQue = "EXC"
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
    BloqueaTexto
End Sub


'================================
Private Sub cmdPrimero_Click()
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
    
Private Sub Botones()

    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdSalir.Enabled = True
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = False
    cmdPrimero.Enabled = False
    cmdultimo.Enabled = False

End Sub



'================================
Sub inicio()
'================================
 adoC.Requery
    If adoC.RecordCount > 0 Then
        If mQue <> "ALT" Then
            adoC.MoveFirst
            ActualizaFormulario
        End If
    Else
        LimpiaBoxes
    End If
    
    cmdIncluir.Enabled = True
    cmdSalir.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    
    
    If adoC.RecordCount = 0 Then
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
        cmdAnterior.Enabled = False
        cmdProximo.Enabled = False
        cmdPrimero.Enabled = False
        cmdultimo.Enabled = False
    Else
        cmdAlterar.Enabled = True
        cmdExcluir.Enabled = True
        cmdAnterior.Enabled = True
        cmdProximo.Enabled = True
        cmdPrimero.Enabled = True
        cmdultimo.Enabled = True
    End If
        

End Sub


'================================
Private Sub LimpiaBoxes()
'================================
  'CERA LOS BOXES
Dim ni As Byte
For ni = 0 To 13
    txtFields(ni).Text = ""
Next
For ni = 0 To 2
    Label1(ni).Caption = ""
Next
End Sub


Private Sub ActualizaFormulario()
txtFields(0).Text = adoC(0)
BuscaSocio
txtFields(1).Text = adoC(1)
BuscaComercio
txtFields(2).Text = adoC(2)
txtFields(3).Text = adoC(3)
txtFields(4).Text = adoC(11)
BuscaMoneda
txtFields(5).Text = Format(adoC(4), "#,#0.00")
txtFields(6).Text = adoC(7)
Label2.Caption = Format(adoC(4) * adoC(7), "#,#0.00")
txtFields(7).Text = adoC(8)
Label3.Caption = Format(adoC(4) * (adoC(7) - adoC(8)), "#,#0.00")
txtFields(8).Text = Format(adoC(9), "#,#0.00")
Label4.Caption = Format(adoC(4) * adoC(7) + adoC(9), "#,#0.00")
txtFields(9).Text = Format(adoC(10), "#,#0.00")
Label5.Caption = Format(adoC(4) * adoC(7) + adoC(9) - adoC(10), "#,#0.00")
txtFields(10).Text = Format(adoC(12), "#,#0.00")
Label6.Caption = Format(adoC(12) * adoC(7), "#,#0.00")
txtFields(11).Text = Format(adoC(13), "#,#0.00")
txtFields(12).Text = adoC(5)
txtFields(13).Text = adoC(6)
txtFields(14).Text = "" & adoC(15)   'tipo
txtFields(15).Text = "" & adoC(16)  'funcionario

'si es una devulucion muestra un cartel por que
If adoC(15) = 4 Then 'es una devolucion
    If adoD.State = adStateOpen Then adoD.Close
    adoD.Open "select * FROM tbl_accion WHERE acc_NroIdent ='" & adoC(2) & "';", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not adoD.RecordCount = 1 Then
        MsgBox "No Encuentro la Anulación"
    Else
        MsgBox "Este conforme fue anulado el " & adoD(4) & vbCrLf & _
            " por " & adoD(3) & vbCrLf & _
            " Motivo: " & adoD(2)
    End If
   If adoD.State = adStateOpen Then adoD.Close
   Set adoD = Nothing
End If
End Sub


Private Sub ActualizaCampos()

End Sub


Private Sub txtFields_Change(Index As Integer)
txtFields(Index).SelStart = 0
txtFields(Index).SelLength = Len(txtFields(Index))
End Sub


'================================
Private Function ValidaCampos() As Boolean
'================================
'If Nro.Text = "" Or Not IsNumeric(Nro.Text) Then
'If Not IsDate(FechaIng.Text) Then
'If Direc.Text = "" Or IsNull(Direc.Text) Then
End Function


Private Sub BloqueaTexto()
Dim ni As Byte
    For ni = 0 To 13
        txtFields(ni).Locked = True
    Next ni
End Sub


Private Sub DesbloqueaTexto()
Dim ni As Byte
    For ni = 0 To 13
        txtFields(ni).Locked = False
    Next ni
End Sub


Private Sub BuscaSocio()
    cSOC.vlNroSoc = CLng(txtFields(0).Text)
    cSOC.mfBuscaSocio
    Label1(0).Caption = Left(cSOC.vsApellido & " " & cSOC.vsNombre, 30)

End Sub


Private Sub BuscaComercio()
    Label1(1).Caption = cCom.BuscaComercio2(CLng(txtFields(1).Text))
End Sub


Private Sub BuscaMoneda()
Select Case txtFields(4).Text
    Case "P"
        Label1(2).Caption = "Pesos"
    Case "D"
        Label1(2).Caption = "Dólares"
    Case "R"
        Label1(2).Caption = "Reales"
    Case "A"
        Label1(2).Caption = "Pesos Arg."
    Case "U"
        Label1(2).Caption = "U.Reaj."
        
End Select
End Sub
