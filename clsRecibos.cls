VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents adoMome As ADODB.Recordset
Attribute adoMome.VB_VarHelpID = -1
Dim WithEvents adoM2 As ADODB.Recordset
Attribute adoM2.VB_VarHelpID = -1
Dim adoRecib As New ADODB.Recordset
Dim cDEP As clsDepend      'clase dependientes
Dim cTC As clsTCambio       'clase Tasa Cambio
Dim cADO As clsAdo

'PARAMETROS
Public vlNumRecibo As Long
Public vnsSocAct As Single
Public vnsSocHon As Single
Public vnsSocCop As Single
Public vnsAyuda As Single
Public vdMes As Date
Public vsDetalle As String

'tasas de cambio
Private vsTCD As Single
Private vsTCR As Single
Private vsTCA As Single
Private vsTCU As Single


'VARIABLES PARA RECIBOS
Private vbCobrador As Byte
Private vlSocio As Long
Private vbCateg As Byte
Private vsNombre As String
Private vsDirec  As String
Private vnsTotal As Single
Private vbAyuda As Boolean
Private vsCuota As Single
Private vsAyuda As Single
Private vsCreditosP As Single
Private vsCreditosME As Single
Private vsCarniceria As Single
Private vsValesP As Single
Private vsValesME As Single





'=================================
Public Function m0fTomaParametros() As Boolean
'=================================
    m0fTomaParametros = False
    Set adoMome = New ADODB.Recordset
    Set adoMome.ActiveConnection = adoconn
    adoMome.Open "SELECT * FROM TBL_Parametros", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    adoMome.MoveFirst
    vlNumRecibo = adoMome("PRM_NroRecibo") + 1
    vnsSocAct = adoMome("PRM_SocAct")
    vnsSocHon = adoMome("PRM_SocHon")
    vnsSocCop = adoMome("PRM_SocCop")
    vnsAyuda = adoMome("PRM_Ayuda")
    vdMes = CDate(adoMome("prm_PrspHst") & "/" & _
        Right(adoMome("prm_prspst"), 2) & "/" & _
        Left(adoMome("prm_prspst"), 4))
    vsDetalle = adoMome("prm_mensaje")
    adoMome.Close
    Set adoMome = Nothing
    
    'toma las tasas de cambio
    Dim cTC As New clsTCambio
    If vsTCD = 0 Then vsTCD = cTC.mfDevuelveCambio("D", vdMes)
    If vsTCR = 0 Then vsTCR = cTC.mfDevuelveCambio("R", vdMes)
    If vsTCA = 0 Then vsTCA = cTC.mfDevuelveCambio("A", vdMes)
    If vsTCU = 0 Then vsTCU = cTC.mfDevuelveCambio("U", vdMes)
    Set cTC = Nothing
    
    m0fTomaParametros = True
End Function



'=================================================
Public Function m1fCargaLosRegistrosAImprimir(dMome As Date, _
    nCobrador As Integer, bSitLab As Byte) As Boolean
'=================================================
'
    m1fCargaLosRegistrosAImprimir = False
    Set adoMome = New ADODB.Recordset
    
    adoMome.CursorLocation = adUseClient
    'Abre Ordenes y su adjunta SOCIOS
    'dEJA 2 CAMPOS VACIOS DNOMB Y DCI
    adoMome.Open "Select *,space(20) as DNomb,space(11) as DCI From TBL_ORDENES AS T1,TBL_SOCIOS AS T2" & _
    " WHERE T2.NroSoc = T1.Ord_NroSoc;", _
    adoconn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    If adoMome.RecordCount < 1 Then
        Exit Function
    End If
    Set adoM2 = adoMome     'PARA QUE NO SE BORREN REGISTROS DE LA TABLA
    
    'Borra los registros que no corresponden
    adoM2.MoveFirst
    Do While Not adoM2.EOF
        'Si ya se pagaron.
        If Not adoM2.Fields("Ord_Cerro") = CDate(0) Then
            adoM2.Delete
        'Vence en presupuestos futuros
        ElseIf adoM2.Fields("Ord_FVto") > dMome Then
            adoM2.Delete
        'No es de este cobrador
        ElseIf Not adoM2.Fields("Cobrador") = nCobrador Then
            adoM2.Delete
        'No es del tipo laboral
        ElseIf Not adoM2.Fields("codSitLab") = bSitLab Then
            adoM2.Delete
        End If
        adoM2.MoveNext
    Loop
    If adoM2.RecordCount < 1 Then
    Exit Function
    End If
    'LLENA DATOS DEPENDIENTES
    Dim cDEP As New clsDepend
    cDEP.mfAbre2TablaDepend
    adoM2.MoveFirst
    Do While Not adoM2.EOF
        cDEP.vlNroSoc = adoM2.Fields("ORD_NroSoc")
        cDEP.vlDepNum = adoM2.Fields("ORD_Depend")
        If cDEP.mfBusca2Depend Then
            adoM2.Fields("DNomb").Value = cDEP.vsDepNmb
            adoM2.Fields("dci").Value = cDEP.vsDepDoc
        End If
        adoM2.MoveNext
    Loop
    Set cDEP = Nothing
    m1fCargaLosRegistrosAImprimir = True
End Function


'=================================
Public Function m2fImprimeRecibos()
'=================================
Dim nSocio As Long


m2sCreaTablaVirtual
vnsTotal = 0

'1 Llena el adoMome
adoM2.Sort = "ord_nrosoc"
adoM2.MoveFirst
m2fInicializa
m2fTomaDatos
nSocio = adoM2.Fields("ord_nrosoc")
Do While Not adoM2.EOF
    If Not adoM2.Fields("ORD_NroSoc") = nSocio Then
        m2fCreaRegistroEnAdoMome
        m2fInicializa
        m2fTomaDatos
        m2fAumentaTotales
    Else
        m2fAumentaTotales
    End If
    nSocio = adoM2.Fields("ord_nrosoc")
    adoM2.MoveNext
Loop
m2fCreaRegistroEnAdoMome
'FIN: llena el adoMome


' Un datagrid para depuracion
Set fjRecibo.DataGrid1.DataSource = adoMome

'2 Graba todos los registros en tbl_RecibosEmitidos
Dim cADO As New clsAdo
cADO.msSalvaUnAdo adoMome, adoconn, "tbl_RecibosEmitidos"

'3 GuardaParamtros: NoRecibo
m2sGuardaParametros


'4 Muestra el data Report
'drRecibo.Sections(1).Controls(33).Caption = vsDetalle
drRecibo.Sections(1).Controls(33).Caption = "Hola Esto es todo lo que se ve desde el nuevo " & _
    " para saber que todo lo que se dice es una nueva d a a la nde"
Set drRecibo.DataSource = adoMome
drRecibo.Sections(1).Controls(1).DataField = "recibo"
drRecibo.Sections(1).Controls(2).DataField = "socio"
drRecibo.Sections(1).Controls(3).DataField = "categoria"
drRecibo.Sections(1).Controls(4).DataField = "nombre"
drRecibo.Sections(1).Controls(5).DataField = "direccion"
drRecibo.Sections(1).Controls(6).DataField = "totcuota"
drRecibo.Sections(1).Controls(7).DataField = "totayuda"
drRecibo.Sections(1).Controls(8).DataField = "totcreditoP"
drRecibo.Sections(1).Controls(9).DataField = "totcreditome"
drRecibo.Sections(1).Controls(10).DataField = "totcarniceria"
drRecibo.Sections(1).Controls(11).DataField = "totvalesP"
drRecibo.Sections(1).Controls(12).DataField = "totvalesME"
drRecibo.Sections(1).Controls(13).DataField = "total"
drRecibo.Sections(1).Controls(14).DataField = "mes"
drRecibo.Sections(1).Controls(31).DataField = "cobrador"

drRecibo.Sections(1).Controls(17).DataField = "recibo"
drRecibo.Sections(1).Controls(18).DataField = "socio"
drRecibo.Sections(1).Controls(19).DataField = "categoria"
drRecibo.Sections(1).Controls(20).DataField = "nombre"
drRecibo.Sections(1).Controls(21).DataField = "direccion"
drRecibo.Sections(1).Controls(22).DataField = "totcuota"
drRecibo.Sections(1).Controls(23).DataField = "totayuda"
drRecibo.Sections(1).Controls(24).DataField = "totcreditoP"
drRecibo.Sections(1).Controls(25).DataField = "totcreditome"
drRecibo.Sections(1).Controls(26).DataField = "totcarniceria"
drRecibo.Sections(1).Controls(27).DataField = "totvalesP"
drRecibo.Sections(1).Controls(28).DataField = "totvalesME"
drRecibo.Sections(1).Controls(29).DataField = "total"
drRecibo.Sections(1).Controls(30).DataField = "mes"
drRecibo.Sections(1).Controls(32).DataField = "cobrador"

drRecibo.Show
End Function


'=================================
Public Sub m2sGuardaParametros()
'=================================
    Set adoM2 = New ADODB.Recordset
    Set adoM2.ActiveConnection = adoconn
    adoM2.Open "SELECT * FROM TBL_Parametros", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    adoM2.MoveFirst
    adoM2("PRM_NroRecibo") = vlNumRecibo
    adoM2.Update
    adoM2.Close
    Set adoM2 = Nothing
End Sub

'=================================
Private Function m2fTomaDatos()
'=================================
vbCobrador = adoM2.Fields("cobrador")
vlSocio = adoM2.Fields("ord_NroSoc")
vbCateg = adoM2.Fields("CodCatSoc")
vsNombre = adoM2.Fields("apellido") & " " & adoM2.Fields("nombre")
vsDirec = adoM2.Fields("direccion") & " " & adoM2.Fields("direccion")
vbAyuda = adoM2.Fields("ayuda")
End Function

'=================================
Private Function m2fInicializa()
'=================================
vbCobrador = 0
vlSocio = 0
vbCateg = 0
vsNombre = ""
vsDirec = ""
vnsTotal = 0
vsCuota = 0
vsAyuda = 0
vsCreditosP = 0
vsCreditosME = 0
vsCarniceria = 0
vsValesP = 0
vsValesME = 0
End Function

'=================================
Private Function m2fCreaRegistroEnAdoMome()
'=================================
adoMome.AddNew
adoMome.Fields("Cobrador") = vbCobrador
adoMome.Fields("Socio") = vlSocio
Select Case vbCateg
    Case 1
        adoMome.Fields("Categoria") = "Activo    "
    Case 2
        adoMome.Fields("Categoria") = "Honorario "
    Case 3
        adoMome.Fields("Categoria") = "Cooperador"
    Case Else
        adoMome.Fields("Categoria") = "Desconoc. "
End Select
adoMome.Fields("Nombre") = vsNombre
adoMome.Fields("Direccion") = vsDirec
adoMome.Fields("Mes") = Right(vdMes, 6)
adoMome.Fields("Recibo") = vlNumRecibo
vlNumRecibo = vlNumRecibo + 1
adoMome.Fields("Total") = vnsTotal
adoMome.Fields("TotCuota") = vsCuota
adoMome.Fields("TotAyuda") = vsAyuda
adoMome.Fields("TotCreditoP") = vsCreditosP
adoMome.Fields("TotCreditoME") = vsCreditosME
adoMome.Fields("TotCarniceria") = vsCarniceria
adoMome.Fields("TotValesP") = vsValesP
adoMome.Fields("TotValesME") = vsValesME
adoMome.Update
End Function


'=================================
Private Function m2fAumentaTotales()
'=================================
'Aumenta los distintos totales
Dim sMomeP As Single
Dim sMomeME As Single
Dim lmome As Long

'toma el nro orden por si es cuota o ayuda
lmome = adoM2.Fields("ord_nroorden")

'valor de la orden
sMomeP = adoM2.Fields("ord_cuota") + _
    adoM2.Fields("ord_recarg") + _
    adoM2.Fields("ord_entcta")
If adoM2.Fields("ord_mon") = "P" Then     'es en pesos
    sMomeME = 0
Else                                    'es en moneda extran
    sMomeME = sMomeP * m2fCotizac(adoM2.Fields("ord_mon"))
    sMomeP = 0
End If


Select Case lmome
    Case 0  'cuota
        vsCuota = vsCuota + sMomeP + sMomeME        'P y ME solo por las dudas
    Case 1  'ayuda
        vsAyuda = vsAyuda + sMomeP + sMomeME        'P y ME solo por las dudas
    Case 115    'carniceria
        vsCarniceria = vsCarniceria + sMomeP + sMomeME
    Case 190    'vales
        vsValesP = vsValesP + sMomeP
        vsValesME = vsValesME * sMomeME
    Case Else
        vsCreditosP = vsCreditosP + sMomeP
        vsCreditosME = vsCreditosME + sMomeME
End Select
vnsTotal = vnsTotal + sMomeP + sMomeME
End Function

'=================================
Private Function m2fCotizac(sPrm As String) As Single
'=================================
Select Case sPrm
    Case "D"
        m2fCotizac = vsTCD
    Case "A"
          m2fCotizac = vsTCA
      
    Case "R"
        m2fCotizac = vsTCR

    Case "U"
        m2fCotizac = vsTCU

    Case Else
        m2fCotizac = 1
End Select
End Function



'=================================
Private Sub m2sCreaTablaVirtual()
'=================================
Set adoMome = New ADODB.Recordset

adoMome.Fields.Append "Cobrador", adChar, 2
adoMome.Fields.Append "Socio", adChar, 6
adoMome.Fields.Append "Categoria", adChar, 12
adoMome.Fields.Append "Nombre", adChar, 50
adoMome.Fields.Append "Direccion", adChar, 50
adoMome.Fields.Append "Mes", adChar, 15
adoMome.Fields.Append "Total", adSingle
adoMome.Fields.Append "Recibo", adInteger, 4
adoMome.Fields.Append "TotCuota", adSingle
adoMome.Fields.Append "TotAyuda", adSingle
adoMome.Fields.Append "TotCreditoP", adSingle
adoMome.Fields.Append "TotCreditoME", adSingle
adoMome.Fields.Append "TotCarniceria", adSingle
adoMome.Fields.Append "TotValesP", adSingle
adoMome.Fields.Append "TotValesME", adSingle


adoMome.CursorType = adOpenDynamic
adoMome.LockType = adLockOptimistic
adoMome.Open
End Sub




'=================================
Public Function mfInvierteMes(ptFecha As String) As String
'=================================
Dim mpFecha As Date

mpFecha = CDate(ptFecha)
mfInvierteMes = Month(mpFecha) & "/" & _
    Day(mpFecha) & "/" & Year(mpFecha)
End Function


'=================================
Public Sub mfGuardaNroRecibo()
'=================================
On Error GoTo merrGNR
    If adoRecib.State = adStateOpen Then adoRecib.Close
    adoRecib.Open "SELECT * FROM tbl_Parametros;", adoconn, adOpenDynamic, adLockOptimistic
    adoRecib.MoveFirst
    adoRecib("PRM_NroRecibo") = vplNroRecibo
    adoRecib.Update
    adoRecib.Close
    Set adoRecib = Nothing
    Exit Sub
merrGNR:
MsgBox "Error 39578: Al guardar Nro Recibo " & Err.Description & "  " & Err.Number
End Sub

'=================================
Public Sub mfTomaNroRecibo()
'=================================
On Error GoTo merrGNR1
    If adoRecib.State = adStateOpen Then adoRecib.Close
    adoRecib.Open "SELECT * FROM tbl_Parametros;", adoconn, adOpenDynamic, adLockOptimistic
    adoRecib.MoveFirst
     vplNroRecibo = adoRecib("PRM_NroRecibo")
    adoRecib.Close
    Set adoRecib = Nothing
    Exit Sub
merrGNR1:
MsgBox "Error 39579: Al guardar Nro Recibo " & Err.Description & "  " & Err.Number
End Sub

'=================================
Public Sub mfTomaYGuardaNroRecibo()
'=================================
On Error GoTo merrGNR1
    If adoRecib.State = adStateOpen Then adoRecib.Close
    adoRecib.Open "SELECT * FROM tbl_Parametros;", adoconn, adOpenDynamic, adLockOptimistic
    adoRecib.MoveFirst
    vplNroRecibo = adoRecib("PRM_NroRecibo")
    adoRecib("PRM_NroRecibo") = vplNroRecibo + 1
    adoRecib.Update
    adoRecib.Close
    Set adoRecib = Nothing
    Exit Sub
merrGNR1:
MsgBox "Error 39580: Al Leer/Guardar Nro Recibo " & Err.Description & "  " & Err.Number
End Sub


