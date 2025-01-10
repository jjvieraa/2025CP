VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjCobros 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Cobros"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "fjCobros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelTodo 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sel.Todo"
      Height          =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtHasta 
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Text            =   "0"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtECta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdECta 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ent.Cta."
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
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
   Begin VB.TextBox txtSocio 
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
      TabIndex        =   4
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "nComerc"
         Caption         =   "Comercio"
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
         DataField       =   "vencim"
         Caption         =   "Vencimiento"
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
         DataField       =   "noorden"
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
      BeginProperty Column03 
         DataField       =   "valorp"
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
         DataField       =   "cuota"
         Caption         =   "Plan"
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
      BeginProperty Column05 
         DataField       =   "valorme"
         Caption         =   "MExt"
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
         DataField       =   "moned"
         Caption         =   "Mon"
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
      BeginProperty Column07 
         DataField       =   "comercio"
         Caption         =   "NoCom"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hasta mes: mmaaaa  (todo = 0)"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No.Socio:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "fjCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' En adoM coloca los registros encontrados en cORD.adoM2
Dim adoM As New ADODB.Recordset
' El adoT es creado en FormLoad con el campo m1, es vacìo unicamente para
' conectar el datareport drcobros y poder ver 2 recibos
Dim adoT As New ADODB.Recordset

Dim cClie As New clsSocios
'Dim cOrd As New clsOrdenes
'Dim cPag As New clsPagos
Dim cRecib As New clsRecibos
Dim cComercio As New clsComercios

Dim sTot2 As Single     'total ordenes
Dim sTot1 As Single     'total seleccionado
Dim kECta As Boolean

Dim nTipoPago As Integer



Private Sub cmdActualizar_Click()
Dim varBmk As Variant
Dim sMensaje1 As String
Dim sMensaje2 As String
Dim sMom As String
Dim sMom1 As String
Dim sECta As Single

Dim nCuenta As Byte

cmdActualizar.Enabled = False
cmdECta.Enabled = False
Screen.MousePointer = vbHourglass

If sTot2 = 0 Then Exit Sub

'If Not mGlob.cPgs.mfAbrePagos Then Exit Sub


sMensaje1 = ""
sMensaje2 = ""

If kECta Then
    sECta = CSng(txtECta.Text)
End If

'recorre los seleccionados en adoM
nCuenta = 1
cRecib.mfTomaNroRecibo
For Each varBmk In DataGrid1.SelBookmarks
        
        If kECta And sECta < 0 Then
            Exit For
        End If
                
        adoM.Bookmark = varBmk
        
        If kECta Then
            If adoM!valorp > sECta Then
                adoM!valorp = sECta
                adoM.Update
            End If
        End If
        'Actualiza la orden
        nTipoPago = cOrd.mfActualizaLaOrden(adoM!NoOrden, adoM!moned, _
            adoM!valorp, adoM!valorME, adoM!socio, adoM!vencim)
        
        If nTipoPago = 0 Then Exit Sub  'no debe nada
        'Guarda en Pagos
        If mGlob.cPgs.mfGuardaUnPago(adoM!socio, adoM!NoOrden, _
            adoM!valorp, Format(Date, "short date"), _
            vplNroRecibo, _
            adoM!comercio, adoM!moned, adoM!valorME, _
            nTipoPago, adoM!cuota & " " & adoM!vencim, _
            Format(Time, "short time"), CStr(vpnFuncionario)) Then
        End If
        If adoM!NoOrden = 1 Then
            sMom = "Cuota"
        ElseIf adoM!NoOrden = 2 Then
            sMom = "Ayuda"
        Else
            sMom = "Orden " & adoM!NoOrden & " " & Trim(adoM!cuota)
        End If
        If nTipoPago = 5 Then           'paga el valor
        ElseIf nTipoPago = 6 Then       'hace una e cta
            sMom = sMom & " ECta "
        End If
        If Trim(adoM!moned) = "" Then 'es en pesos
            sMensaje2 = sMensaje2 & "  (" & nCuenta & ")  " & _
                sMom & " Vto:" & adoM!vencim & "  $" & Format(adoM!valorp, "#,#0.00") & " " & cComercio.BuscaComercio2(adoM!comercio)
        Else    'es en ME
            sMensaje2 = sMensaje2 & "  (" & nCuenta & ")  " & _
                sMom & " Vto:" & adoM!vencim & "  $" & Format(adoM!valorp, "#,#0.00") & _
                "(U$" & Format(adoM!valorME, "#,#0.00") & ")" & " " & cComercio.BuscaComercio2(adoM!comercio)
        End If
        nCuenta = nCuenta + 1
        sECta = sECta - adoM!valorp
 Next
DoEvents
txtSocio.SetFocus


'Imprime

'Paga siempre en pesos
sMom1 = "P"
'If adoM!moned = "" Then
'    sMom1 = "P"
'Else
'    sMom1 = adoM!moned
'End If
If kECta Then
    sTot1 = CSng(txtECta.Text)
End If
sMensaje1 = "Rivera, " & Date & vbCrLf & _
    "Recibimos de " & Trim(Label2.Caption) & vbCrLf & _
    " Socio No." & txtSocio.Text & "   " & Label8.Caption & _
    vbCrLf & "la cantidad de $ " & _
    Format(sTot1, "#,#0.00") & "  (" & mfPalabras(sTot1, sMom1) & ")"

Set drCobros.DataSource = adoT          'para que muestre 2 recibos
drCobros.Sections(3).Controls(1).Caption = sMensaje1
drCobros.Sections(3).Controls(2).Caption = "por los siguientes pagos: " & sMensaje2
drCobros.Sections(3).Controls(3).Caption = "Recibo No. " & vplNroRecibo
    
Screen.MousePointer = vbDefault


drCobros.Show

'actualiza el numero de recibo
vplNroRecibo = vplNroRecibo + 1
cRecib.mfGuardaNroRecibo

End Sub




Private Sub cmdECta_Click()
txtECta.Visible = True
cmdActualizar.Enabled = True
cmdECta.Enabled = False
cmdECta.Enabled = False
kECta = True
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSelTodo_Click()
While Not adoM.EOF
      DataGrid1.SelBookmarks.Add adoM.Bookmark
      DataGrid1_Click
   adoM.MoveNext
Wend

End Sub

Private Sub Form_Load()
cmdActualizar.Enabled = False
cmdECta.Enabled = False
txtECta.Visible = False
cmdSelTodo.Enabled = False
kECta = False


'Crea un ado momentaneo para generar 1(2) registros en la impresión
Set adoT.ActiveConnection = adoConn
Set adoT = New ADODB.Recordset
adoT.Fields.Append "m1", adChar, 2
adoT.Open
adoT.AddNew
'adoT.AddNew
adoT.Update

If Not cComercio.mfAbreTablaComercios Then
    MsgBox "Error 3323: Al abrir tabla COmercios."
    Unload Me
End If

cClie.msInicia
If Not cClie.mfAbreTablaSociosOrdenSocio Then
    MsgBox "ERROR 956: Abriendo tabla Socios " & Err.Description
    Exit Sub
End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
Set cClie = Nothing
Set cOrd = Nothing
'Set cPag = Nothing
Set cRecib = Nothing
If adoM.State = adStateOpen Then adoM.Close
Set adoM = Nothing
If adoT.State = adStateOpen Then adoT.Close
Set adoT = Nothing
Set cComercio = Nothing
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

Private Sub txtHasta_Validate(Cancel As Boolean)
If Not IsNumeric(txtHasta.Text) Then Cancel = True
End Sub

Private Sub txtSocio_GotFocus()
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
End Sub

Private Sub txtSocio_KeyDown(KeyCode As Integer, Shift As Integer)
           If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrSocAlf9
                fjMuestraTabla.Show
            ElseIf KeyCode = 114 Then   'F3
                vpMuestraTabla = kMstrSocPorNC2    'Por No Cobro
                fjMuestraTabla.Show
            End If

End Sub

Private Sub txtSocio_Validate(Cancel As Boolean)
If Not IsNumeric(txtSocio.Text) Then Cancel = True
End Sub


Private Sub cmdVer_Click()
        '1 Pide No de cliente: busca si existe
        cClie.vlNroSoc = CLng(txtSocio.Text)
        If Not cClie.mfBuscaSocio Then
                Label2.Caption = "No encontrado"
                txtSocio.SetFocus
                Exit Sub
        Else
                Label2.Caption = cClie.vsApellido & "  " & cClie.vsNombre
                Label8.Caption = "No.Cobro: " & cClie.vsNroCob & "   Cat:" & cClie.vlCodCatSoc
                Label3.Caption = "Espere..."
        End If
         
        '2 Busca la deuda en glob
        mGlob.cOrd.vlNroSoc = CLng(txtSocio.Text)
        If mGlob.cOrd.fBuscaOrdenes2UnSocio = False Then
                Label3.Caption = ""
                txtSocio.SetFocus
                Exit Sub
        End If
        
        ' No tiene deuda
        If mGlob.cOrd.adoOrdenes.RecordCount < 1 Then
                MsgBox "Sin deuda", vbExclamation, "Resultado de la búsqueda"
                Exit Sub
        End If
        
        
        'Set fjMome.DataGrid1.DataSource = cOrd.adoOrdenes
        'fjMome.Show
        'Exit Sub
        mGlob.cOrd.msPreparaOrdenesAPagarEnAdoM2 (1)
        sTot2 = 0
        ' No tiene deuda
        If mGlob.cOrd.adoM2.RecordCount < 1 Then
                MsgBox "Sin deuda", vbExclamation, "Resultado de la búsqueda"
                Exit Sub
        End If
        
        Set adoM = cOrd.adoM2
        'cOrd.adoM2.Close
        'Set cOrd.adoM2 = Nothing
        
        'si es hasta un mes
        If Not txtHasta = 0 Then
                Dim tHasta
                tHasta = vpnPrspHst + 1 & "/" & Left(txtHasta.Text, 2) & _
                    "/" & Right(txtHasta.Text, 4)
                adoM.Filter = "vencim < #" & tHasta & "#"
        End If
        '3) Suma la deuda
        adoM.MoveFirst
        Do While Not adoM.EOF
                sTot2 = sTot2 + adoM!valorp
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

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
    '3 Va mostrando el total de lo que se selecciona
    Dim varBmk As Variant
    sTot1 = 0
    
    For Each varBmk In DataGrid1.SelBookmarks
            adoM.Bookmark = varBmk
            sTot1 = sTot1 + CSng(adoM!valorp)
    Next
    Label5.Caption = "Selecc:      " & Format(sTot1, "#,#0.00")
    Label6.Caption = "No Selec:    " & Format(sTot2 - sTot1, "#,#0.00")
End Sub

