VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjPagoAComerc 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Pagos A Comercios"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "fjPagosAComercios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelTodo 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sel.Todo"
      Height          =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtECta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdECta 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ent.Cta."
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdVer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtComerc 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "F2=ALF F3=N COB"
      Top             =   120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   12648447
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
         DataField       =   "D_Orden"
         Caption         =   "Orden"
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
         DataField       =   "d_Clie"
         Caption         =   "Cliente"
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
         DataField       =   "d_FVto"
         Caption         =   "Vto"
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
         DataField       =   "d_Haber"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "d_Debe"
         Caption         =   "Cuotas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "d_dscto"
         Caption         =   "Desc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "otro"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,#0.00"
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
            Alignment       =   1
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha Pago:"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No.Comerc:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "fjPagoAComerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoM As New ADODB.Recordset
Dim adoS As New ADODB.Recordset

Dim cCom As New clsComercios

Dim sTot2 As Single     'total ordenes
Dim sTot1 As Single     'total seleccionado
Dim kECta As Boolean

Dim nTipoPago As Integer



Private Sub Form_Load()
cmdActualizar.Enabled = False
cmdECta.Enabled = False
txtECta.Visible = False
cmdSelTodo.Enabled = False
kECta = False

cCom.mfAbreTablaComercios

If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT *, d_haber-d_debe-d_dscto as otro FROM tbl_DeudComerc " & _
    "WHERE NOT d_Recibo=6 " & _
    "AND NOT d_FVto > #" & mfInvierteMes(CStr(fjComercios.Text1.Text)) & "# " & _
    "AND NOT d_cerro AND abs(d_haber - d_dscto - d_debe) >2 ORDER BY d_orden;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

'abre la tbl_salidas
If adoS.State = adStateOpen Then adoS.Close
adoS.Open "SELECT * FROM tbl_Salidas;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

End Sub


Private Sub txtComerc_GotFocus()
Set DataGrid1.DataSource = Nothing
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label6.Caption = ""
Label8.Caption = ""
cmdVer.Enabled = True
cmdActualizar.Enabled = False
cmdSelTodo.Enabled = False
cmdECta.Enabled = False
DataGrid1.Refresh
txtECta.Visible = False
kECta = False
sTot1 = 0
sTot2 = 0
End Sub



Private Sub txtComerc_KeyDown(KeyCode As Integer, Shift As Integer)
           If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrComerc3
                fjMuestraTabla.Show
            End If

End Sub



Private Sub txtComerc_Validate(Cancel As Boolean)
If Not IsNumeric(txtComerc.Text) Then Cancel = True
End Sub


Private Sub cmdVer_Click()

'1 Pide No de comercio: busca si existe
Label2.Caption = cCom.BuscaComercio2(CLng(txtComerc.Text))

If Label2.Caption = "No encontrado..." Then
    Exit Sub
End If

'2 Despliega su deuda
adoM.Filter = "d_Comercio=" & CLng(txtComerc.Text)
adoM.Requery
If adoM.RecordCount = 0 Then
    MsgBox "Sin saldo a cobrar"
    Exit Sub
Else
    Label8.Caption = adoM.RecordCount & " regs."
End If

'3) Suma la deuda
adoM.MoveFirst
Do While Not adoM.EOF
    sTot2 = sTot2 + adoM!d_haber - adoM!d_debe - adoM!d_dscto
    adoM.MoveNext
Loop

Set DataGrid1.DataSource = adoM
Label4.Caption = "Total: " & Format(sTot2, "#,#0.00")
sTot1 = 0
Label3.Caption = "Utilice Ctrl para varios registros"
cmdVer.Enabled = False
cmdActualizar.Enabled = True
cmdSelTodo.Enabled = True
cmdECta.Enabled = True



End Sub



Private Sub cmdActualizar_Click()
Dim varBmk As Variant
Dim sECta As Single
Dim sMome As Single


'no hay nada que actualizar
If sTot2 = 0 Then Exit Sub
'no hay nada seleccionado
If sTot1 = 0 Then Exit Sub

'que tenga fecha
If Not mfEsFecha(txtFecha.Text) Then
    txtFecha.SetFocus
    Exit Sub
End If



'si es entrega a cuenta
If kECta Then
    sECta = CSng(txtECta.Text)
End If

'recorre los seleccionados en adoM
For Each varBmk In DataGrid1.SelBookmarks
        'es entrega a cta y ya se completó
        If kECta And sECta < 0 Then
            Exit For
        End If
                
        adoM.Bookmark = varBmk
        'el valor que se está saldando
        sMome = adoM!d_haber - adoM!d_debe - adoM!d_dscto
        'es entrega a cta
        If kECta Then
            If sMome > sECta Then
                sMome = sECta
            End If
        End If
        'Actualiza el registro
        If adoM!d_recibo = 6 Then       'ojo no  hay de estos
            adoM!d_haber = adoM!d_haber + sMome
        Else
            adoM!d_debe = adoM!d_debe + sMome
        End If
        'cancelo la deuda
        If adoM!d_haber = adoM!d_debe Then
            adoM!d_Cerro = True
        End If
        adoM.Update
        
        'Guarda en Salidas
        AgregoRegSalida (sMome)
        sECta = sECta - sMome
 Next
txtComerc.SetFocus



End Sub


Private Sub AgregoRegSalida(sprm As Single)
adoS.AddNew
adoS!s_fecha = CDate(txtFecha.Text)
adoS!s_valor = sprm
adoS!s_Nume = CLng(txtComerc.Text)
adoS!s_Deta = adoM!d_Orden
adoS!s_FNume = 1
adoS!s_func = vpnFuncionario
adoS!s_fdia = Format(Date, "short date")
adoS!s_fHora = Format(Time, "short time")
adoS.Update
End Sub




Private Sub cmdECta_Click()
txtECta.Visible = True
cmdECta.Enabled = False
kECta = True
End Sub


Private Sub cmdSelTodo_Click()
While Not adoM.EOF
      DataGrid1.SelBookmarks.Add adoM.Bookmark
      DataGrid1_Click
   adoM.MoveNext
Wend

End Sub






Private Sub txtECta_GotFocus()
txtECta.SelStart = 0
txtECta.SelLength = Len(txtECta.Text)
End Sub

Private Sub txtECta_Validate(Cancel As Boolean)
If Not IsNumeric(txtECta.Text) Then Cancel = True
If CSng(txtECta.Text) > sTot1 Then
    MsgBox "Entrega mayor que Selección"
    Cancel = True
End If
End Sub








Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
'3 Va mostrando el total de lo que se selecciona
Dim varBmk As Variant
sTot1 = 0
For Each varBmk In DataGrid1.SelBookmarks
        adoM.Bookmark = varBmk
        sTot1 = sTot1 + adoM!d_haber - adoM!d_debe - adoM!d_dscto
 Next
Label5.Caption = "Selecc:      " & Format(sTot1, "#,#0.00")
Label6.Caption = "No Selec:    " & Format(sTot2 - sTot1, "#,#0.00")

End Sub


Private Sub Form_Unload(Cancel As Integer)

If adoM.State = adStateOpen Then adoM.Close
Set adoM = Nothing
If adoS.State = adStateOpen Then adoS.Close
Set adoS = Nothing

Set cCom = Nothing
End Sub

Private Sub txtFecha_Change()
        If Len(txtFecha.Text) = 2 Then
            txtFecha.Text = txtFecha.Text & "/"
            txtFecha.SelStart = 3
        ElseIf Len(txtFecha.Text) = 5 Then
            txtFecha.Text = txtFecha.Text & "/"
            txtFecha.SelStart = 6
        End If

End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
If Not IsDate(txtFecha.Text) Then Cancel = True
End Sub
