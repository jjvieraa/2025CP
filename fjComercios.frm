VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fjComercios 
   BackColor       =   &H00800000&
   Caption         =   "Comercios"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4035
   Icon            =   "fjComercios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdFIn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdVer 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ver"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdPago 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pago"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdListaUno 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lista Uno"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdListaTodo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lista Todo"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd2Cierre 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cierre"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Nro Comercio"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Nro Comercio"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1440
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   3735
   End
End
Attribute VB_Name = "fjComercios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim adoP As New ADODB.Recordset
    Dim adoM As New ADODB.Recordset
    Dim adoQ As New ADODB.Recordset
    Dim adoCmd As New ADODB.Command
    Dim cTC As New clsTCambio
    Dim dFechaInicioEj As Date
    Dim dFechaFinEj As Date
    
    
    Dim sTCDolar As Single
    Dim sTCReal As Single
    Dim sTCAus As Single

    Dim kTipoListado As Byte
    Const kListaTodos = 1
    Const kListaUno = 2
    Const kPagos = 3




'==================================================
Private Sub Command1_Click()
'==================================================
    Screen.MousePointer = vbHourglass
MomeImportaDeudaAtrtasadaAComercios
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFIn_Click()
Unload Me
End Sub

Private Sub cmdPago_Click()
DeshabilitaBotones
Label2.Caption = "Hasta mes 10/mm/aaaa: "
Label2.Visible = True
Text1.Visible = True
Text1.SetFocus
kTipoListado = kPagos
cmdVer.Visible = True
End Sub
'==================================================
Private Sub cmdVer_Click()
'==================================================
Select Case kTipoListado
    Case kListaTodos
        ListaTodo
    Case kListaUno
        ListaUno
    Case kPagos
        fjPagoAComerc.Show
End Select
HabilitaBotones
End Sub

'==================================================
Private Sub Form_Load()
'==================================================
   If vpbMesOperac = 1 Then
            dFechaInicioEj = CDate(vpnPrspHst + 1 & "/12/" & vpnAñoOperac - 1)
    Else
            dFechaInicioEj = CDate(vpnPrspHst + 1 & "/" & vpbMesOperac - 1 & "/" & vpnAñoOperac)
    End If
    dFechaFinEj = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
    HabilitaBotones
End Sub



'==================================================
Private Sub mCierraTodo()
'==================================================
    Mensaje24 ""
    PB.Visible = False
    If adoP.State = adStateOpen Then adoP.Close
    Set adoP = Nothing
    If adoM.State = adStateOpen Then adoM.Close
    Set adoM = Nothing
    If adoQ.State = adStateOpen Then adoQ.Close
    Set adoQ = Nothing
    Screen.MousePointer = vbDefault
    Text1.Visible = False
    Label2.Visible = False

End Sub
'==================================================
Private Sub mfGuardaComercio(lCom As Long, dFecha As Date, _
    sHaber As Single, lRecibo As Long, sDebe As Single)
'==================================================
    adoM.AddNew
    adoM("d_comercio") = lCom
    adoM("d_fecha") = dFecha
    adoM("d_Debe") = sDebe
    adoM("d_haber") = sHaber
    adoM("d_recibo") = lRecibo
    adoM("d_func") = vpnFuncionario
    adoM("d_fdia") = Format(Date, "short date")
    adoM("d_Fhora") = Format(Time, "short time")
End Sub

'==================================================
Private Sub Mensaje24(sprm As String)
'==================================================
    Label1.Caption = sprm
    Label1.Refresh
End Sub

'==================================================
Private Sub cmdListaUno_Click()
'==================================================
DeshabilitaBotones
Label2.Caption = "Hasta mes 10/mm/aaaa: "
Label2.Visible = True
Text1.Visible = True
kTipoListado = kListaUno
cmdVer.Visible = True
Text1.SetFocus
Label3.Visible = True
Text2.Visible = True

End Sub
Private Sub DeshabilitaBotones()
cmd2Cierre.Enabled = False
cmdListaTodo.Enabled = False
cmdListaUno.Enabled = False
cmdPago.Enabled = False
cmdVer.Enabled = True
End Sub

Private Sub HabilitaBotones()
cmd2Cierre.Enabled = True
cmdListaTodo.Enabled = True
cmdListaUno.Enabled = True
cmdPago.Enabled = True
Text1.Visible = False
Label2.Visible = False
cmdVer.Visible = False
PB.Visible = False
Label3.Visible = False
Text2.Visible = False
cmdVer.Enabled = False
End Sub
'==================================================
Private Sub ListaUno()
'==================================================
     Dim cn As New ADODB.Connection
    Dim sM As String
    Dim nM As Long
    
    

    Screen.MousePointer = vbHourglass
    Mensaje24 "Archivos..."
    
    
    '1) CARGA LOS REGISTROS
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Provider = "MSDATASHAPE"
    cn.Open "dsn=jimmy"
      sM = "SELECT *, NombCom & ' ' & Codigo & space(15) & ' Telf.' & tel as B3, d_haber - d_dscto - d_debe as b4,d_haber - d_debe as b5  FROM tbl_DeudComerc as T1 " & _
            "INNER JOIN tbl_comercios as T2 " & _
            "ON T2.codigo = T1.D_Comercio " & _
            "WHERE d_Comercio=" & CLng(Text2.Text) & _
           " AND NOT d_FVto > #" & _
            mfInvierteMes(CStr(Text1.Text)) & "# AND " & _
            "d_Recibo < 6  AND NOT d_cerro AND abs(d_haber - d_dscto - d_debe) >2 ORDER BY d_Comercio, d_Orden, d_FVto;"
    
   
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "SHAPE {" & sM & "}  AS cm1 COMPUTE cm1 BY 'b3'"
        .Execute
    End With
    
    If adoM.State = adStateOpen Then adoM.Close
    With adoM
        .ActiveConnection = cn
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open adoCmd
    End With
    'Set adoQ = adoM(0).Value
    If adoM.RecordCount < 0 Then
        MsgBox "Sin Registros"
        GoTo final2
    End If
    'Set fjMome.fdg1.DataSource = adoM
    'Set fjMome.DataGrid1.DataSource = adoM
    'fjMome.Show
    'Exit Sub
    Mensaje24 "Preparando Listado..."
    'adoM.Filter = "cm1.B4 > 5"
    
    
    '9) Muestra el informe de datos
    dr2Comercios.Hide
    dr2Comercios.Caption = "Resumen de Ordenes por Comercios al " & Text1.Text
    dr2Comercios.Title = "Resumen de Ordenes por Comercios" & vbCrLf & "al " & Text1.Text
    Set dr2Comercios.DataSource = adoM
    dr2Comercios.DataMember = ""
    
    'GRUPO
    dr2Comercios.Sections(3).Controls(1).DataMember = ""
    dr2Comercios.Sections(3).Controls(1).DataField = "b3"
    
    'DETALLES
    dr2Comercios.Sections(4).Controls(1).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(1).DataField = "d_Orden"
    dr2Comercios.Sections(4).Controls(2).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(2).DataField = "d_FVto"
    dr2Comercios.Sections(4).Controls(3).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(3).DataField = "b5"
    dr2Comercios.Sections(4).Controls(4).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(4).DataField = "d_Dscto"
    dr2Comercios.Sections(4).Controls(5).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(5).DataField = "d_NoC"
    dr2Comercios.Sections(4).Controls(6).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(6).DataField = "d_plan"
    dr2Comercios.Sections(4).Controls(7).DataMember = "cm1"
    dr2Comercios.Sections(4).Controls(7).DataField = "B4"           'COMERCIO
    
    'PIE DE GRUPO
    'dr2comercios.Sections(5).Controls(1).DataMember = ""
    'dr2comercios.Sections(5).Controls(1).DataField = "i_sNombre"
    dr2Comercios.Sections(5).Controls(1).DataMember = "cm1"
    dr2Comercios.Sections(5).Controls(1).DataField = "b5"
    dr2Comercios.Sections(5).Controls(2).DataMember = "cm1"
    dr2Comercios.Sections(5).Controls(2).DataField = "d_Dscto"
    dr2Comercios.Sections(5).Controls(3).DataMember = "cm1"
    dr2Comercios.Sections(5).Controls(3).DataField = "b4"
    'dr2comercios.Sections(5).Controls(3).DataMember = "tbl_Info01"
    'dr2comercios.Sections(5).Controls(3).DataField = "ord_meCuota"
     
     'PIE DE INFORME
'dr2Comercios.Sections(7).Controls(1).DataMember = "cm1"
 '   dr2Comercios.Sections(7).Controls(1).DataField = "d_haber"
 '   dr2Comercios.Sections(7).Controls(2).DataMember = "cm1"
 '   dr2Comercios.Sections(7).Controls(2).DataField = "d_dscto"
 '   dr2Comercios.Sections(7).Controls(3).DataMember = "cm1"
 '   dr2Comercios.Sections(7).Controls(3).DataField = "b4"

    'dr2comercios.Sections(7).Controls(2).DataMember = "tbl_Info01"
    'dr2comercios.Sections(7).Controls(2).DataField = "ord_meCuota"
    
    dr2Comercios.Refresh
    Screen.MousePointer = vbDefault
    DoEvents
    dr2Comercios.Show    'Set fjMome.fdg1.DataSource = adoM
    'Set fjMome.DataGrid1.DataSource = adoM
    'fjMome.Show
    'Exit Sub
final2:
mCierraTodo
   

End Sub
'==================================================
Private Sub cmdListaTodo_Click()
'==================================================
DeshabilitaBotones
Label2.Caption = "Hasta mes 10/mm/aaaa: "
Label2.Visible = True
Text1.Visible = True
kTipoListado = kListaTodos
cmdVer.Visible = True
Text1.SetFocus

    
    
   
End Sub


'==================================================
Private Sub ListaTodo()
'==================================================
    
    
    
    
    Dim cn As New ADODB.Connection
    Dim sM As String
    Dim nM As Long
    
    

    Screen.MousePointer = vbHourglass
    Mensaje24 "Archivos..."
    
    
    '1) CARGA LOS REGISTROS
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Provider = "MSDATASHAPE"
    cn.Open "dsn=jimmy"
    sM = "SELECT *, NombCom & ' ' & Codigo & space(15) & ' Telf.' & tel as B3, d_haber - d_dscto as b4,d_haber - d_debe as b5  FROM tbl_DeudComerc as T1 " & _
            "INNER JOIN tbl_comercios as T2 " & _
            "ON T2.codigo = T1.D_Comercio " & _
            "WHERE NOT T1.d_FVto > #" & _
            mfInvierteMes(CStr(Text1.Text)) & "# AND " & _
            "NOT (d_comercio = 86 OR d_comercio = 100 OR d_comercio = 115 OR d_comercio = 190) AND " & _
            "t1.d_Recibo < 6 AND NOT d_cerro AND abs(d_haber - d_dscto - d_debe) >2 ORDER BY t1.d_Comercio, t1.d_Orden, t1.d_FVto;"

    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "SHAPE {" & sM & "}  AS cm1 COMPUTE cm1 BY 'b3'"
        .Execute
    End With
    
    If adoM.State = adStateOpen Then adoM.Close
    With adoM
        .ActiveConnection = cn
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open adoCmd
    End With
    'Set adoQ = adoM(0).Value
    If adoM.RecordCount < 0 Then
        MsgBox "Sin Registros"
        GoTo final2
    End If
    'Set fjMome.fdg1.DataSource = adoM
    'Set fjMome.DataGrid1.DataSource = adoM
    'fjMome.Show
    'Exit Sub
    Mensaje24 "Preparando Listado..."
    
    '9) Muestra el informe de datos
    drComercios.Hide
    drComercios.Caption = "Resumen de Ordenes por Comercios al " & Text1.Text
    drComercios.Title = "Resumen de Ordenes por Comercios" & vbCrLf & "al " & Text1.Text
    Set drComercios.DataSource = adoM
    drComercios.DataMember = ""
    
    'GRUPO
    drComercios.Sections(3).Controls(1).DataMember = ""
    drComercios.Sections(3).Controls(1).DataField = "b3"
    
    'DETALLES
    drComercios.Sections(4).Controls(1).DataMember = "cm1"
    drComercios.Sections(4).Controls(1).DataField = "d_Orden"
    drComercios.Sections(4).Controls(2).DataMember = "cm1"
    drComercios.Sections(4).Controls(2).DataField = "d_FVto"
    drComercios.Sections(4).Controls(3).DataMember = "cm1"
    drComercios.Sections(4).Controls(3).DataField = "b5"
    drComercios.Sections(4).Controls(4).DataMember = "cm1"
    drComercios.Sections(4).Controls(4).DataField = "d_Dscto"
    drComercios.Sections(4).Controls(5).DataMember = "cm1"
    drComercios.Sections(4).Controls(5).DataField = "d_NoC"
    drComercios.Sections(4).Controls(6).DataMember = "cm1"
    drComercios.Sections(4).Controls(6).DataField = "d_plan"
    drComercios.Sections(4).Controls(7).DataMember = "cm1"
    drComercios.Sections(4).Controls(7).DataField = "B4"           'COMERCIO
    
    'PIE DE GRUPO
    'drComercios.Sections(5).Controls(1).DataMember = ""
    'drComercios.Sections(5).Controls(1).DataField = "i_sNombre"
    drComercios.Sections(5).Controls(1).DataMember = "cm1"
    drComercios.Sections(5).Controls(1).DataField = "b5"
    drComercios.Sections(5).Controls(2).DataMember = "cm1"
    drComercios.Sections(5).Controls(2).DataField = "d_Dscto"
    drComercios.Sections(5).Controls(3).DataMember = "cm1"
    drComercios.Sections(5).Controls(3).DataField = "b4"
    'drComercios.Sections(5).Controls(3).DataMember = "tbl_Info01"
    'drComercios.Sections(5).Controls(3).DataField = "ord_meCuota"
     
     'PIE DE INFORME
     drComercios.Sections(7).Controls(1).DataMember = "cm1"
    drComercios.Sections(7).Controls(1).DataField = "b5"
    drComercios.Sections(7).Controls(2).DataMember = "cm1"
    drComercios.Sections(7).Controls(2).DataField = "d_dscto"
    drComercios.Sections(7).Controls(3).DataMember = "cm1"
    drComercios.Sections(7).Controls(3).DataField = "b4"

    'drComercios.Sections(7).Controls(2).DataMember = "tbl_Info01"
    'drComercios.Sections(7).Controls(2).DataField = "ord_meCuota"
    
    drComercios.Refresh
    Screen.MousePointer = vbDefault
    DoEvents
    drComercios.Show
    'Set fjMome.fdg1.DataSource = adoM
    'Set fjMome.DataGrid1.DataSource = adoM
    'fjMome.Show
    'Exit Sub
final2:
mCierraTodo

End Sub



'==================================================
Private Sub msLlenaUnRegDeInfo01()
'==================================================
Dim sPorc As Single
Dim sValor As Single
Dim sDcto As Single
adoQ.AddNew
adoQ("ord_NroOrden") = adoP("ord_NroOrden")
adoQ("ord_NroSoc") = adoP("ord_NroSoc")
adoQ("ord_NroCom") = adoP("ord_NroCom")
adoQ("Ord_FEmis") = adoP("ord_femis")
adoM.MoveFirst
adoM.Find ("Codigo =" & adoP("ord_NroCom"))
If Not adoM.EOF Then
    adoQ("i_sNombre") = adoM("NombCOm")
    sPorc = 0 + adoM("desc")
Else
    adoQ("i_sNombre") = "Desconocido"
    sPorc = 0
End If
sValor = adoP("ord_cuota") * adoP("ord_plan")
adoQ("Ord_Cuota") = sValor
sDcto = sValor * sPorc / 100        ' el descuento sobre el valor
adoQ("ord_recarg") = sDcto      'en recargo queda el descuento
adoQ("ord_entCta") = sValor - sDcto 'en ent cta queda el total
If Not adoP("ord_Mon") = "P" Then
    adoQ("i_sM1") = Format(adoP("ord_MECuota") * adoP("ord_plan"), "#,#0.00")
End If

adoQ.Update
End Sub


'==================================================
Private Sub ms2LlenaUnRegDeInfo01(sPorc As Single)
'==================================================
Dim sValor As Single
Dim sDcto As Single
adoQ.AddNew
adoQ("ord_NroOrden") = adoP("ord_NroOrden")
adoQ("ord_NroSoc") = adoP("ord_NroSoc")
adoQ("ord_NroCom") = adoP("ord_NroCom")
adoQ("Ord_FEmis") = adoP("ord_femis")
sValor = adoP("ord_cuota") * adoP("ord_plan")
adoQ("Ord_Cuota") = sValor
sDcto = sValor * sPorc / 100        ' el descuento sobre el valor
adoQ("ord_recarg") = sDcto      'en recargo queda el descuento
adoQ("ord_entCta") = sValor - sDcto 'en ent cta queda el total
If Not adoP("ord_Mon") = "P" Then
    adoQ("i_sM1") = Format(adoP("ord_MECuota") * adoP("ord_plan"), "#,#0.00")
End If

adoQ.Update
End Sub

Private Sub Text1_change()
        If Len(Text1.Text) = 2 Then
            Text1.Text = Text1.Text & "/"
            Text1.SelStart = 3
        ElseIf Len(Text1.Text) = 5 Then
            Text1.Text = Text1.Text & "/"
            Text1.SelStart = 6
        End If

End Sub

'==================================================
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)  'jv
'==================================================
    Select Case kTipoListado
        Case kPagos
        Case kListaTodos
        Case kListaUno
            If KeyCode = 113 Then   'F2
                vpMuestraTabla = kMstrComerc2
                fjMuestraTabla.Show
            End If
    End Select
End Sub


'==================================================
Private Sub Text1_Validate(Cancel As Boolean)
'==================================================
        If Not IsDate(Text1.Text) Then Cancel = True
End Sub
'==================================================
Private Sub Text2_Validate(Cancel As Boolean)
'==================================================
        If Not IsNumeric(Text2.Text) Then Cancel = True
End Sub



'==================================================
Private Sub cmdMome_Click()
'==================================================
'es solo un ejemplo de un listado sin uso de tabla
    Dim cn As New ADODB.Connection
    Dim sM As String
    Dim nM As Long
    
    
    PB.Visible = True
    'Screen.MousePointer = vbHourglass
    Mensaje24 "Archivos..."
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Provider = "MSDATASHAPE"
    cn.Open "dsn=jimmy"
    sM = "SELECT ord_NroCom,ord_NroOrden, ord_NroSoc, ord_Femis, " & _
            "ord_Cuota*ord_plan as St1, ord_mon, ord_mecuota*ord_plan as St2, " & _
            "ord_cuota, ord_recarg, ord_entcta," & _
            "ord_tipo,space(3) as xx1, T2.Codigo, T2.NombCom,t2.Desc  FROM tbl_Ordenes as T1 " & _
            "INNER JOIN tbl_comercios as T2 " & _
            "ON T2.codigo = T1.ord_nrocom " & _
            "WHERE ord_FEmis BETWEEN #" & _
            mfInvierteMes(CStr(dFechaInicioEj)) & "# AND #" & _
            mfInvierteMes(CStr(dFechaFinEj)) & "# AND " & _
            "NOT ord_NroCom = 0 AND " & _
            "NOT ord_tipo = 4  ORDER BY ord_NroCom,ord_NroOrden;"

    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "SHAPE {" & sM & "}  AS cm1 COMPUTE cm1 BY 'ord_NroCom'"
        .Execute
    End With
    
    If adoM.State = adStateOpen Then adoM.Close

    With adoM
        .ActiveConnection = cn
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open adoCmd
    End With
    Set adoQ = adoM(0).Value
    
    'Set fjMome.fdg1.DataSource = adoM
    'Set fjMome.DataGrid1.DataSource = adoM
    'fjMome.Show
    'Exit Sub
    
    Mensaje24 "Recorre Ado..."
    PB.Min = 0
    nM = 0
    PB.Max = adoM.RecordCount
  Dim rsvar As Variant
  Dim sPorc As Single
  Dim sDcto As Single
    adoQ.MoveFirst
    Do While Not adoM.EOF
        'Debug.Print adoQ(0), adoQ(1)
        nM = nM + 1
        adoQ("ord_cuota") = adoQ("st1")
        sPorc = 0 + adoQ("desc")
        sDcto = adoQ("st1") * sPorc / 100        ' el descuento sobre el valor
        adoQ("ord_recarg") = sDcto      'en recargo queda el descuento
        adoQ("ord_entCta") = adoQ("st1") - sDcto 'en ent cta queda el total

        If nM Mod 100 = 0 Then PB.Value = nM
        adoM.MoveNext
    Loop
    Set fjMome.fdg1.DataSource = adoM
    Set fjMome.DataGrid1.DataSource = adoQ
    fjMome.Show
    Exit Sub
  
   ' drComercios.Hide
   ' Set drMome.DataSource = adoM
   ' drMome.DataMember = ""
    
    'GRUPO
   ' drMome.Sections(3).Controls(1).DataMember = ""
    'drMome.Sections(3).Controls(1).DataField = "ord_NroCOm"
    
    'DETALLES
    'drMome.Sections(4).Controls(1).DataMember = "cm1"
    'drMome.Sections(4).Controls(1).DataField = adoQ(0).Name
    'drMome.Refresh
 
    'drMome.Show
final2:
mCierraTodo
End Sub



'==================================================
Private Sub cmd2Cierre_Click()
'==================================================
    If vpnNivelFuncionario < kNivel6 Then
        MsgBox "Sin Autorización"
       Exit Sub
    End If
    PB.Visible = True
    DeshabilitaBotones
    Screen.MousePointer = vbHourglass
'0
  Dim sMo As String
  sMo = "Continúa el cierre ?" & vbCrLf & _
            "Si continúa, se realizan las siguientes acciones: " & vbCrLf & _
            "Se toman las ordenes generadas entre " & dFechaInicioEj & _
            " y " & dFechaFinEj & vbCrLf & " y se inserta en tbl_DeudComerc " & _
            "un registro por cada cuota" & vbCrLf
    If MsgBox(sMo, vbYesNo + vbQuestion, "Cierre Mes Comercios!") = vbNo Then
                GoTo termina
          End If
'1) Verifica que mes anterior este cerrado en tbl_accion
    Mensaje24 "7.Verifica Mes Anterior..."
    If Not rCierraMes_VerifMesAnteriorEsteCompleto(6) Then
          If MsgBox("Cancela el cierre ?", vbYesNo + vbQuestion, "Mes Anterior Incompleto!") = vbYes Then
                GoTo termina
          End If
    End If
    
'2) Verifica que ya no se haya cerrado el mes anteriormente
    Dim sM As String
    Mensaje24 "6.Verificando repeticion..."
    If Not rCierraMes_VerifMesEsteCompleto(6) Then
          If MsgBox("Cancela el cierre ?", vbYesNo + vbQuestion, "Repitiendo el Cierre!") = vbYes Then
                GoTo termina
          ElseIf MsgBox("Elimina los registros anteriores (recomendado)?", vbYesNo + vbQuestion, "Repitiendo el Cierre!") = vbYes Then
               sM = "DELETE * FROM tbl_DeudComerc " & _
               "WHERE (d_recibo = 5  OR d_recibo = 6 ) AND " & _
               "d_Fecha =#" & _
                mfInvierteMes(CStr(dFechaFinEj)) & "#;"
                If adoP.State = adStateOpen Then adoP.Close
                adoP.Open sM, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
          End If
    End If

'3) toma la tasa de cambio del dia
Mensaje24 "Las TC deben ser distintas de cero"

Screen.MousePointer = vbDefault
sTCDolar = 0
Do While sTCDolar = 0
    sTCDolar = cTC.mfDevuelveCambio("D", dFechaFinEj)
Loop
sTCReal = 0
Do While sTCReal = 0
    sTCReal = cTC.mfDevuelveCambio("R", dFechaFinEj)
Loop
sTCAus = 0
Do While sTCAus = 0
    sTCAus = cTC.mfDevuelveCambio("A", dFechaFinEj)
Loop
Screen.MousePointer = vbHourglass


'4) Genera el ado: 1 registro por cuota
' y con D_recibo=5
    
    Mensaje24 "5.Genera Ado..."
    
    If adoP.State = adStateOpen Then adoP.Close
    sM = "SELECT * FROM tbl_Ordenes INNER JOIN tbl_comercios " & _
         "ON tbl_comercios.codigo = tbl_ordenes.ord_nrocom " & _
        " WHERE ord_FEmis BETWEEN #" & _
            mfInvierteMes(CStr(dFechaInicioEj)) & "# AND #" & _
            mfInvierteMes(CStr(dFechaFinEj)) & "# AND" & _
            " NOT ord_NroCom = 0 AND" & _
            " NOT ord_tipo = 4;"
    adoP.Open sM, adoconn, adOpenKeyset, adLockOptimistic, adCmdText


    PB.Min = 0
    PB.Max = adoP.RecordCount
    Dim nM As Long
    Dim bM As Byte
    Mensaje24 "4.Recorre Ado..."
   
    nM = 0
    adoCmd.ActiveConnection = adoconn
    adoCmd.CommandType = adCmdText
    adoP.MoveFirst
    Do While Not adoP.EOF
        For bM = 1 To adoP("ord_plan")
            mf2GuardaComercio bM
        Next
        nM = nM + 1
        If nM Mod 100 Then PB.Value = nM
        adoP.MoveNext
    Loop

    
'4) Coloca registros por la cuota mensual
'1 registro por comercio cooperador con recibo=6
    If adoP.State = adStateOpen Then adoP.Close
    adoP.Open "SELECT * FROM tbl_Comercios;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    PB.Min = 0
    PB.Max = adoP.RecordCount
    Mensaje24 "3.Cooperadores..."
    nM = 0
     If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM tbl_DeudComerc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

    adoP.MoveFirst
    Do While Not adoP.EOF
        If Not adoP("Cooperador") Then
            mfGuardaComercio adoP("Codigo"), dFechaFinEj, 0, 6, vpsCuotaSCop
        End If
        nM = nM + 1
        If nM Mod 100 Then PB.Value = nM
        adoP.MoveNext
    Loop
    adoM.UpdateBatch adAffectAllChapters

'5) actualiza las cuotas en ME
'iy marcA las cerradas
    If adoP.State = adStateOpen Then adoP.Close
    adoP.Open "SELECT * FROM tbl_DeudComerc WHERE NOT d_cerro;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    PB.Min = 0
    PB.Max = adoP.RecordCount
    Mensaje24 "2.Moneda Extranj."
    nM = 0
    adoP.MoveFirst
    Do While Not adoP.EOF
        If adoP("d_debe") = adoP("d_haber") Then
            adoP("d_cerro") = True
        ElseIf Not adoP("d_Mone") = "P" Then
            Select Case adoP("d_Mone")
                Case "D"
                    adoP("d_haber") = adoP("d_ValorME") * sTCDolar
                Case "R"
                    adoP("d_haber") = adoP("d_ValorME") * sTCReal
                Case "A"
                    adoP("d_haber") = adoP("d_ValorME") * sTCAus
                Case Else   'es es pesos ???
                    adoP("d_haber") = adoP("d_ValorME")
            End Select
        End If
        nM = nM + 1
        If nM Mod 100 Then PB.Value = nM
        adoP.MoveNext
    Loop
    adoM.UpdateBatch adAffectAllChapters

   
'5) Pone marca en tbl_accion
    Mensaje24 "1.Marca..."
    If Not fColocaMarcaAccion(26, vptMesPresup, "Cierre Comercio", "", "") Then
            MsgBox "Error 24233: No se completó marca de Cierre"
    End If
HabilitaBotones
termina:
    mCierraTodo
End Sub
  

'==================================================
Private Sub mf2GuardaComercio(nPlan As Byte)
'==================================================
Dim dMome As Date
Dim sValor As Single
Dim sValorME As Single
Dim sDesc As Single

'If adoP("ord_NroOrden") = 5467 Then
'    sValor = 0
'End If

If nPlan > 1 Then
    dMome = mfAgregaMesesAFecha(adoP("ord_FVto"), nPlan - 1)
Else
    dMome = adoP("ord_FVto")
End If
If adoP("ord_Mon") = "P" Then
    sValor = adoP("ord_cuota")
    sValorME = 0
Else
    Select Case adoP("ord_Mon")
        Case "D"
            sValor = adoP("ord_MECuota") * sTCDolar
          Case "R"
            sValor = adoP("ord_MECuota") * sTCReal
       Case "A"
            sValor = adoP("ord_MECuota") * sTCAus
       Case Else
            sValor = adoP("ord_mecuota")
  End Select
 sValorME = adoP("ord_mecuota")
End If
sDesc = 0 + adoP("desc")

adoCmd.CommandText = "insert into tbl_DeudComerc " & _
"(d_comercio,d_Clie, d_orden, d_fecha, d_Recibo, d_Debe, d_haber, d_deta, " & _
"d_plan, d_NoC,  d_FVto, d_Mone, d_ValorME, d_Dscto, d_Func, d_FDia, d_fHora) " & _
"values ('" & _
    adoP("ord_NroCom") & "','" & _
    adoP("ord_NroSoc") & "','" & _
    adoP("ord_NroOrden") & "','" & _
    dFechaFinEj & _
    "','5' , '0','" & sValor & "','','" & _
    adoP("ord_plan") & "','" & nPlan & "','" & _
    dMome & "','" & adoP("ord_mon") & "','" & sValorME & _
    "','" & sValor * sDesc / 100 & "','" & vpnFuncionario & _
    "','" & Date & "','" & Time & "')"
adoCmd.Execute
End Sub


'==================================================
Private Sub MomeImportaDeudaAtrtasadaAComercios()
'==================================================
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.Open "SELECT * FROM impComerc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM tbl_DeudComerc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    adoQ.MoveFirst
    Do While Not adoQ.EOF
        mf5GuardaComercio
        adoQ.MoveNext
    Loop
mCierraTodo
End Sub




'==================================================
Private Sub mf5GuardaComercio()
'==================================================
    adoM.AddNew
    adoM("d_comercio") = adoQ("cta_com")
    adoM("d_clie") = adoQ("cta_cli")
    adoM("d_Orden") = adoQ("cta_Nro")
    adoM("d_Fecha") = CDate("01/12/02")
    adoM("d_Recibo") = 5
    adoM("d_Haber") = adoQ("cta_sal")
    If adoQ("CTA_MES") = 1 Then
            adoM("d_FVto") = CDate("10/12/" & adoQ("cta_ano") - 1)
    Else
        adoM("d_FVto") = CDate("10/" & adoQ("cta_mes") - 1 & "/" & adoQ("cta_ano"))
    End If
    adoM("d_NoC") = adoQ("cta_cta")
    adoM("d_Mone") = "P"
    adoM("d_Dscto") = adoQ("CTA_SAL") * adoQ("CTA_POR") / 100
    adoM.Update
End Sub


