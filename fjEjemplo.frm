VERSION 5.00
Begin VB.Form fjEjemplo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TBL_Ordenes"
   ClientHeight    =   6390
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   5775
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   5280
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
         Picture         =   "fjEjemplo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Picture         =   "fjEjemplo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "fjEjemplo.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Picture         =   "fjEjemplo.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   375
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
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
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Cerro"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   28
      Top             =   4540
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_MEPagos"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   26
      Top             =   4220
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_MECuota"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   24
      Top             =   3900
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Mon"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   22
      Top             =   3580
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_EntCta"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   20
      Top             =   3260
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Recarg"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   18
      Top             =   2940
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_CtasPagas"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   16
      Top             =   2620
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Plan"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   14
      Top             =   2300
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_FVto"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_FEmis"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   10
      Top             =   1660
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Cuota"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   1340
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_Depend"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_NroOrden"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   700
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_NroCom"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   380
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ORD_NroSoc"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_Cerro:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   27
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_MEPagos:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   25
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_MECuota:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   23
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_Mon:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   21
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_EntCta:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_Recarg:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   17
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_CtasPagas:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_Plan:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_FVto:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_FEmis:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_Cuota:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_Depend:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_NroOrden:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD_NroCom:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   380
      Width           =   1815
   End
End
Attribute VB_Name = "fjEjemplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoC As New ADODB.Recordset
Dim mQue As String




'=======================================================
Private Sub Form_Load()
'=======================================================
    On Error GoTo mEr333
    
    'la base de datos--------------
    Set adoC.ActiveConnection = adoconn
    If adoC.State = adStateOpen Then adoC.Close
    adoC.Open "select * FROM tbl_Ordenes ORDER BY ORD_nrosoc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
 
    Call inicio
    Call BloqueaTexto
    Call cmdPrimero_Click
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
'Set cCOMR = Nothing
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
    'SSTab1.Tab = 0
    'SSTab1.TabEnabled(2) = False
    Botones
    LimpiaBoxes
    DesbloqueaTexto
    'NroSoc.Enabled = True
    'toma el ultimo No para aumentarlo
    adoC.MoveLast
    'NroSoc.Text = adoC("nrosoc") + 1
    Fech_ing.Text = Date
    NroCob.SetFocus
End Sub

'================================
Private Sub cmdsalir_Click()
'================================
   Unload Me

End Sub


Private Sub cmdAlterar_Click()
    mQue = "ALT"
    DesbloqueaTexto
    'SSTab1.Tab = 0
    'SSTab1.TabEnabled(2) = False
    Botones
    ActualizaFormulario
End Sub


Private Sub cmdCancelar_Click()
    mQue = ""
    BloqueaTexto
    'SSTab1.TabEnabled(2) = True
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
 
End Sub


Private Sub ActualizaFormulario()
txtFields(0).Text = adoC(0)
txtFields(1).Text = adoC(1)
txtFields(2).Text = adoC(2)
txtFields(3).Text = adoC(3)
txtFields(4).Text = adoC(11)
txtFields(5).Text = adoC(4)
txtFields(6).Text = adoC(7)
txtFields(7).Text = adoC(8)
txtFields(8).Text = adoC(9)
txtFields(9).Text = adoC(10)
txtFields(10).Text = adoC(12)
txtFields(11).Text = adoC(13)
txtFields(12).Text = adoC(5)
txtFields(13).Text = adoC(6)
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
Dim nI As Byte
    For nI = 0 To 13
        txtFields(nI).Locked = True
    Next nI
End Sub


Private Sub DesbloqueaTexto()
Dim nI As Byte
    For nI = 0 To 13
        txtFields(nI).Locked = False
    Next nI
End Sub

