VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' El recordset principal es mglob.adoOrdns
Public adoOrdenes As New ADODB.Recordset    ' Auxiliar
Private adoNumOrden As New ADODB.Recordset
Public adoM2 As New ADODB.Recordset
'Private adoB As New ADODB.Recordset
Private sPresupActual As String         'aaaamm
Dim cTC As New clsTCambio
Public vdPrsptoTermina As Date           '10/mm/aaaa

Public vlNroSoc As Long
Public vlNroComerc As Long
Public vlNroDepend As Long
Public vlNroOrden As Long
Public vnsCuota As Single
Public vdFEmis As Date
Public vdFVto As Date
Public vnPlan As Byte
Public vnCtasPaga As Byte
Public vnsEntCta As Single
Public vnsRecarg As Single
Public vsMoneda As String
Public vnsMECuota As Single
Public vnsMEPagos As Single
Public vdCerro As Date
Public vsFunc As String
Public vsFDia As String
Public vsFHora As String




'=====================================================================
Public Sub InicializaVariables()
'=====================================================================
    vlNroSoc = 0
    vlNroComerc = 0
    vlNroDepend = 0
    vnsCuota = 0#
    vdFEmis = CDate("00/00/00")
    vdFVto = CDate("00/00/00")
    vnPlan = 0
    vnCtasPaga = 0
    vnsEntCta = 0#
    vnsRecarg = 0#
    vsMoneda = "P"
    vnsMECuota = 0#
    vnsMEPagos = 0#
    vdCerro = CDate("00/00/00")
End Sub


'=====================================================================
Public Sub msInicia()
'=====================================================================
    Set adoOrdenes = New ADODB.Recordset
    Set adoOrdenes.ActiveConnection = adoConn
    Set adoM2 = New ADODB.Recordset
    Set adoM2.ActiveConnection = adoConn
End Sub

'=====================================================================
Public Sub msInicia2()
'=====================================================================
    Set adoOrdenes = New ADODB.Recordset
    Set adoOrdenes.ActiveConnection = adoConn
    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    adoOrdenes.Open "SELECT * FROM tbl_Ordenes;", adoConn, adOpenDynamic, adLockOptimistic
End Sub



Public Sub msAbreTablaOrdenesGlobal()
    Set mGlob.adoOrdns = New ADODB.Recordset
    mGlob.adoOrdns.Open "SELECT * FROM tbl_Ordenes;", adoConn, adOpenDynamic, adLockOptimistic
End Sub

Public Sub msOrdenaTablaOrdenesGlobalPorNroOrden()
    mGlob.adoOrdns.Sort = "ord_NroOrden"
End Sub

'=====================================================================
Public Sub msTermina()
'=====================================================================
    If adoOrdenes.State = adStateOpen Then
        adoOrdenes.Close
    End If
    Set adoOrdenes = Nothing
End Sub


'=====================================================================
Public Function fBuscaOrdenesUnSocio() As Boolean
'=====================================================================
    On Error GoTo merr5233
    
    Dim sMom As String
    'Dim dMM As Date
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT ORD_NroSoc,ORD_NroCom , ORD_NroOrden," & _
        "ORD_DEPEND, ORD_Cuota, ORD_FEmis,  " & _
        "ORD_FVTO, cdate('01/01/1900')  as Otro, " & _
        "ORD_Plan, ORD_CtasPagas, ORD_EntCta, ORD_Recarg, ORD_Mon, ORD_MECuota, " & _
        "ORD_MEPagos, ord_Auto FROM TBL_Ordenes " & _
        "WHERE ORD_NroSoc =" & _
        CStr(vlNroSoc) & " and year(ord_cerro) < 1901 " & _
        "ORDER BY ORD_NroSoc;"
    adoOrdenes.Open sMom, adoConn, adOpenStatic, adLockOptimistic
    If adoOrdenes.RecordCount < 1 Then
        fBuscaOrdenesUnSocio = False
        Exit Function
    End If
    'Set fjMome.DataGrid1.DataSource = adoOrdenes
    'fjMome.Show
    'MsgBox "Probando"
    Call msColocaFechaVtoEnOtro
    
   fBuscaOrdenesUnSocio = True
    Exit Function
    
merr5233:
    MsgBox "ERROR 5233: " & Err.Description & " EN: " & Err.Number
    fBuscaOrdenesUnSocio = False
End Function




'=====================================================================
Public Function fBuscaTodasLasOrdenesUnSocio() As Boolean
'=====================================================================
    On Error GoTo merr5244
    
    Dim sMom As String
    'Dim dMM As Date
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT *,cdate('01/01/1900')  as Otro FROM TBL_Ordenes " & _
        "WHERE ORD_NroSoc =" & _
        CStr(vlNroSoc) & ";"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount < 1 Then
        fBuscaTodasLasOrdenesUnSocio = False
        Exit Function
    End If
    Call msColocaFechaVtoEnOtro
    
   fBuscaTodasLasOrdenesUnSocio = True
    Exit Function
    
merr5244:
    MsgBox "ERROR 5244: " & Err.Description & " EN: " & Err.Number
    fBuscaTodasLasOrdenesUnSocio = False
End Function




'=====================================================================
Public Function fBuscaOrdenes2UnSocio() As Boolean
'=====================================================================
    On Error GoTo merr5230
    
    Dim sMom As String
    'Dim dMM As Date
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT *, cdate('01/01/1900')  as Otro " & _
        "FROM TBL_Ordenes INNER JOIN tbl_comercios " & _
        "ON tbl_comercios.codigo = tbl_ordenes.ord_nrocom " & _
        "WHERE ORD_NroSoc =" & CStr(vlNroSoc) & _
        " and year(ord_cerro) < 1901;"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount < 1 Then
        fBuscaOrdenes2UnSocio = False
        Exit Function
    End If
    Call msColocaFechaVtoEnOtro
    
   fBuscaOrdenes2UnSocio = True
    Exit Function
    
merr5230:
    MsgBox "ERROR 5230: " & Err.Description & " EN: " & Err.Number
    fBuscaOrdenes2UnSocio = False
End Function

'=====================================================================
Public Function fBuscaUnaOrden(lPrm As Long) As Boolean
'=====================================================================
    On Error GoTo merr5230
    
    Dim sMom As String
    'Dim dMM As Date
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT * FROM tbl_ordenes " & _
        "WHERE ORD_NroOrden =" & CStr(lPrm) & ";"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount <> 1 Then
        msGuardaFalla 0, lPrm, Date, "No encuentra orden"
        MsgBox "Error 72321a: Orden " & lPrm & " duplicada"
        fBuscaUnaOrden = False
        Exit Function
    End If
    
   fBuscaUnaOrden = True
    Exit Function
    
merr5230:
    MsgBox "ERROR 72321a: " & Err.Description & " EN: " & Err.Number
    fBuscaUnaOrden = False
End Function

'=====================================================================
Public Function fBuscaUnaOrden2(lPrm As Long) As Boolean
'=====================================================================
    'previo necesita: msInicia2
    
    On Error GoTo merr5231
    
    adoOrdenes.MoveFirst
    adoOrdenes.Find ("ORD_NroOrden =" & CStr(lPrm))
    If adoOrdenes.EOF Then
        MsgBox "Error 72321d: Orden " & lPrm & " duplicada"
        fBuscaUnaOrden2 = False
        Exit Function
    End If
    
   fBuscaUnaOrden2 = True
    Exit Function
    
merr5231:
    MsgBox "ERROR 7321d: " & Err.Description & " EN: " & Err.Number
    fBuscaUnaOrden2 = False
End Function

'=====================================================================
Public Function fBuscaUnaCuotaOAyuda(lSoc As Long, lOrd As Long, _
fVto As Date) As Boolean
'=====================================================================
    On Error GoTo merr5231
    
    Dim sMom As String
    'Dim dMM As Date
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT * FROM tbl_ordenes " & _
        "WHERE ORD_NroSoc =" & CStr(lSoc) & _
        " AND ord_NroOrden =" & CStr(lOrd) & _
        " AND ord_fVto =#" & mfInvierteMes(CStr(fVto)) & "#;"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount = 0 Then
        MsgBox "Error 72321b: Orden " & lOrd & " No encontrada"
        fBuscaUnaCuotaOAyuda = False
        Exit Function
    ElseIf Not adoOrdenes.RecordCount = 1 Then
        MsgBox "Error 72321b: Orden " & lOrd & " duplicada"
        fBuscaUnaCuotaOAyuda = False
        Exit Function
    End If
    
   fBuscaUnaCuotaOAyuda = True
    Exit Function
    
merr5231:
    MsgBox "ERROR 72321b: " & Err.Description & " EN: " & Err.Number
    fBuscaUnaCuotaOAyuda = False
End Function

'=====================================================================
Public Sub msColocaFechaVtoEnOtro()
'=====================================================================
    Dim nMM As Integer
    'Si tiene ent.cta. lo toma como pagos:
    'recorre el ado
    With adoOrdenes
    .MoveFirst
        
        Do While Not .EOF
            'coloca la fecha de vencimiento en el campo otro
            Debug.Print !ord_NroOrden & "  " & !otro
            If !ord_ctasPagas > 0 Then
                !otro = mfAgregaMesesAFecha(!ord_FVto, !ord_ctasPagas)
            Else
                !otro = !ord_FVto
            End If
            If Not !ord_EntCta = 0 Or Not !ord_Recarg = 0 Then
                nMM = Fix((!ord_EntCta - !ord_Recarg) / !ord_cuota)
                !otro = mfAgregaMesesAFecha(!otro, nMM)
            End If
            .MoveNext
        Loop
    End With

 End Sub
'=====================================================================
Public Sub ms5ColocaFechaVtoEnOtro()
'=====================================================================
    Dim nMM As Integer
    'Si tiene ent.cta. lo toma como pagos:
    'recorre el ado
    With adoOrdenes
    .MoveFirst
        
        Do While Not .EOF
            !otro = !ord_FVto
            .MoveNext
        Loop
    End With

 End Sub
'=====================================================================
Public Sub msPreparaOrdenesAPagarEnAdoM2(nPrm As Byte)
'=====================================================================
    'nprm=5 no coloca ordenes vencidas
    '1)CALCULA EN PESOS LAS ORDENES QUE DEBE PAGAR EN EL SIG.MES
    '2)LAS PONE EN UNA TABLA
    Dim vvNoCuota As Byte

   
    
    '1)Toma el mes del Presupuesto Actual
    Dim nMes As Byte
    Dim nAnio As Integer
    'OJO abre y cierra el adoM2
    sPresupActual = fTomaPresupActual
   
    nMes = CInt(Right(sPresupActual, 2))
    nAnio = CInt(Left(sPresupActual, 4))
    
    vdPrsptoTermina = CDate("10/" & nMes & "/" & nAnio)
    
    '2) Crea la tabla virtual en el adoM2
    Call mfMOAPCreaTablaVirtual(nPrm)
    
    '3) recorre el ado
    Dim nM1 As Single
    Dim nCoc As Integer
    Dim nRes As Single
    
    With adoOrdenes
        .MoveFirst
        Do While Not .EOF
                    Debug.Print !ord_NroOrden
            'Para el nPrm=5 no coloca ordenes vencidas
                    'Numero de pagos
                    'Sin recargo o entCta
                    If !ord_Recarg = !ord_EntCta Then
                        nCoc = !ORD_PLAN - !ord_ctasPagas
                    'hay entcTA
                    ElseIf !ord_Recarg < !ord_EntCta Then
                        'EN PESOS
                        If !ord_Mon = "P" Then
                            nM1 = (!ORD_PLAN - !ord_ctasPagas) * !ord_cuota + !ord_Recarg - !ord_EntCta
                            'por si hubo entregas o recargos
                            nCoc = Fix(nM1 / !ord_cuota)    'PARTE ENTERA
                            nRes = nM1 - nCoc * !ord_cuota     'RESTO
                            If Abs(nRes) > 5 Then nCoc = nCoc + 1
                        Else
                            nM1 = (!ORD_PLAN - !ord_ctasPagas) * !ord_mecuota + !ord_Recarg - !ord_EntCta
                            'por si hubo entregas o recargos
                            nCoc = Fix(nM1 / !ord_mecuota)
                            nRes = nM1 - nCoc * !ord_mecuota     'RESTO
                            If !ord_Recarg < !ord_EntCta Then
                                If Abs(nRes) > 1 Then nCoc = nCoc + 1
                            End If
                        End If
                    Else    'recargo > entcta
                        nCoc = !ORD_PLAN - !ord_ctasPagas
  
                    End If
                    'RECORRE LAS CUOTAS
                    Dim nTot As Integer     'por el total de cuotas
                    nTot = nCoc
                    Do While nCoc > 0
                        '3.1) Realiza los calculos=
                        '3.2) Los guarda en la tabla
                        Call mfMOAP31(nCoc, nTot, nPrm)
                        '3.3) Realiza la suma
                        'sMome = sMome + fSumaOrdenes1
                        nCoc = nCoc - 1
                    Loop
                    .MoveNext
        Loop
    End With
    
  '4) Ordena la tabla
If adoM2.RecordCount > 0 Then Me.msOrdena1
  
 
End Sub

'=====================================================================
Public Sub ms7PreparaOrdenesAPagarEnAdoM2(nPrm As Byte)
'=====================================================================
    'nprm=1 pesos
    'nprm=2 dolares
    
    'oRDENA POR CUOTAS, SIN IMPORTAR PAGOS, ENT.CTA. O RECARG
    
    Dim vvNoCuota As Byte
    '1)Toma el mes del Presupuesto Actual
    Dim nMes As Byte
    Dim nAnio As Integer
    'OJO abre y cierra el adoM2
    sPresupActual = fTomaPresupActual
    nMes = CInt(Right(sPresupActual, 2))
    nAnio = CInt(Left(sPresupActual, 4))
    vdPrsptoTermina = CDate("10/" & nMes & "/" & nAnio)
    
    '2) Crea la tabla virtual en el adoM2
    Call mfMOAPCreaTablaVirtual(0)
    
    '3) recorre el ado
    Dim nM1 As Single
    Dim nCoc As Integer
    Dim nRes As Long
    
    With adoOrdenes
        .MoveFirst
        Do While Not .EOF
            nCoc = !ORD_PLAN
            'RECORRE LAS CUOTAS
            Do While nCoc > 0
                        '3.1) Realiza los calculos=
                        '3.2) Los guarda en la tabla
                        Call mf7MOAP31(nCoc, !ORD_PLAN, nPrm)
                        '3.3) Realiza la suma
                        'sMome = sMome + fSumaOrdenes1
                        nCoc = nCoc - 1
            Loop
            .MoveNext
        Loop
    End With
    
  '4) Ordena la tabla
  Me.msOrdena1
  
 
End Sub
'=====================================================================
Public Sub msMuestraAdoM2(dg As DataGrid)
'=====================================================================
 '5) Muestra la tabla en un datagrid
  Set dg.DataSource = adoM2
  
  '6) Organiza las columnas
  dg.Columns(0).Caption = "Cmr"
  dg.Columns(0).Width = 500
  dg.Columns(1).Caption = "Ord"
  dg.Columns(1).Width = 500
  dg.Columns(2).Caption = "Depend"
  dg.Columns(2).Width = 700
  dg.Columns(2).Alignment = dbgRight
  dg.Columns(3).Caption = "Val.ME"
  dg.Columns(3).Width = 900
  dg.Columns(3).Alignment = dbgRight
  dg.Columns(3).NumberFormat = "#,#"
  dg.Columns(4).Caption = "Vencim."
  dg.Columns(4).Width = 1100
  dg.Columns(5).Caption = "Mnd"
  dg.Columns(5).Width = 500
  dg.Columns(6).Caption = "Valor"
  dg.Columns(6).Width = 900
  dg.Columns(6).Alignment = dbgRight
  dg.Columns(6).NumberFormat = "#,#"
  dg.Columns(7).Visible = False
  'adoM2.MoveFirst           'va al comienzo
  'SendKeys "{HOME}"
  'dg.Row = 0
  'dg.SetFocus
  'dg.RowBookmark = 1
  'dg.FirstRow
End Sub


'=====================================================================
Private Sub mfMOAP31(bNroCuota As Integer, nTot As Integer, nPrm As Byte)
'=====================================================================
    Dim mValor As Single
    Dim mValorME As Single
    Dim dMM As Date
    Dim bAgregaRegRecarg As Boolean
    
    bAgregaRegRecarg = False
    'Dim nPag As Integer
    
    '1)
    Debug.Print adoOrdenes!ord_NroOrden
    adoM2.AddNew
    adoM2!comercio = adoOrdenes!ORD_NroCom
    adoM2!NoOrden = adoOrdenes!ord_NroOrden
    adoM2!NoDepend = adoOrdenes!ORD_DEPEND
    adoM2!vencim = adoOrdenes!otro
    adoM2!socio = adoOrdenes!ord_nrosoc
    adoM2!nAuto = adoOrdenes!ord_Auto
    
    With adoOrdenes
    '2)Cuota
    If bNroCuota = nTot Then       'ES LA PRIMERA
        mValor = !ord_Recarg - !ord_EntCta
        If mValor = 0 Then
            mValor = !ord_cuota
            mValorME = !ord_mecuota
        ElseIf mValor < 0 Then       'ent cta
            mValor = (!ORD_PLAN - !ord_ctasPagas - nTot + 1) * !ord_cuota + !ord_Recarg - !ord_EntCta
            mValorME = (!ORD_PLAN - !ord_ctasPagas - nTot + 1) * !ord_mecuota * cTC.mfDevuelveCambio(!ord_Mon, Date) + !ord_Recarg - !ord_EntCta
            'nPag = CInt((!ord_EntCta - !ord_Recarg) / !ord_cuota)
            'mValor = (!ORD_PLAN - !ord_ctasPagas - nTot - nPag + 1) * !ord_cuota
            'mValorME = (!ORD_PLAN - !ord_ctasPagas - nTot + 1) * !ord_mecuota + !ord_Recarg - !ord_EntCta
        Else        'recargo
            mValor = !ord_cuota
            mValorME = !ord_mecuota
            bAgregaRegRecarg = True
        End If
    Else                        'NO ES LA PRIMERA
        mValor = !ord_cuota
        mValorME = !ord_mecuota
    End If
    
    adoM2!valorp = mValor
    adoM2!cuota = CStr(!ORD_PLAN - bNroCuota + 1) & "/" & CStr(!ORD_PLAN)
    
    '3) Es en moneda extranjera
    If Not !ord_Mon = "P" Then
        adoM2!valorME = mValorME
        adoM2!moned = !ord_Mon
        adoM2!valorp = mValorME * cTC.mfDevuelveCambio(!ord_Mon, Date)
    End If
    
    '4)Es de este presupuesto
     If adoOrdenes!otro > vdPrsptoTermina Then
        adoM2!pr = True
    Else
        adoM2!pr = False
     End If
    If nPrm = 1 Then
        adoM2!ncomerc = "" & Left(Trim(!NombCom), 15)
    End If
    adoM2.Update
    
    
    If bAgregaRegRecarg Then
        adoM2.AddNew
        adoM2!comercio = adoOrdenes!ORD_NroCom
        adoM2!NoOrden = adoOrdenes!ord_NroOrden
        adoM2!NoDepend = adoOrdenes!ORD_DEPEND
        adoM2!vencim = adoOrdenes!otro
        adoM2!socio = adoOrdenes!ord_nrosoc
        adoM2!nAuto = adoOrdenes!ord_Auto
        adoM2!valorp = !ord_Recarg - !ord_EntCta
        adoM2!cuota = "0/" & CStr(!ORD_PLAN)
        
        'Es de este presupuesto
         If adoOrdenes!otro > vdPrsptoTermina Then
            adoM2!pr = True
        Else
            adoM2!pr = False
         End If
         adoM2.Update
  
    End If
    
    '5)Cambia el vencimiento (para el proximo registro)
    'MsgBox Day(!otro)
    'MsgBox Month(!otro)
    'MsgBox Year(!otro)
    If Month(!otro) = 12 Then
        dMM = CDate(Format(Day(!otro) & "/1/" & Year(!otro) + 1, "short date"))
        Else
        dMM = CDate(Format(Day(!otro) & "/" & Month(!otro) + 1 & "/" & Year(!otro), "short date"))
    End If
    !otro = dMM         'lo cambia en el propio adoOrdenes
   End With
    
    
 
End Sub

'=====================================================================
Private Sub mf7MOAP31(bNroCuota As Integer, nTot As Integer, nPrm As Byte)
'=====================================================================
    Dim mValor As Single
    Dim mValorME As Single
    Dim dMM As Date
    
    '1)
    adoM2.AddNew
    adoM2!comercio = adoOrdenes!ORD_NroCom
    adoM2!NoOrden = adoOrdenes!ord_NroOrden
    'adoM2!NoDepend = adoOrdenes!ORD_DEPEND
    adoM2!vencim = adoOrdenes!otro
    adoM2!socio = adoOrdenes!ord_nrosoc
    'adoM2!nAuto = adoOrdenes!ord_Auto
    
    With adoOrdenes
    '2)Cuota
    'es en pesos
    If nPrm = 1 Then
        adoM2!valorp = !ord_cuota
    'es en dolares
    Else
        adoM2!valorp = !ord_mecuota
  
    End If
    'adoM2!cuota = CStr(!ORD_PLAN - bNroCuota + 1) & "/" & CStr(!ORD_PLAN)
    
    '3) Es en moneda extranjera
    'If Not !ord_Mon = "P" Then
    '    adoM2!valorME = mValorME
    '    adoM2!moned = !ord_Mon
    '    adoM2!valorp = mValorME * cTC.mfDevuelveCambio(!ord_Mon, Date)
    'End If
    
    '4)Es de este presupuesto
     'If adoOrdenes!Otro > vdPrsptoTermina Then
     '   adoM2!pr = True
    'Else
     '   adoM2!pr = False
     'End If
    
    '5)Cambia el vencimiento (para el proximo registro)
    'MsgBox Day(!otro)
    'MsgBox Month(!otro)
    'MsgBox Year(!otro)
    If Month(!otro) = 12 Then
        dMM = CDate(Format(Day(!otro) & "/1/" & Year(!otro) + 1, "short date"))
        Else
        dMM = CDate(Format(Day(!otro) & "/" & Month(!otro) + 1 & "/" & Year(!otro), "short date"))
    End If
    !otro = dMM         'lo cambia en el propio adoOrdenes
    'adoM2!ncomerc = "" & Left(Trim(!NombCom), 15)
    End With
    
    '6)
    adoM2.Update
End Sub

'=====================================================================
Private Sub mfMOAPCreaTablaVirtual(nPrm As Byte)
'=====================================================================
Set adoM2 = New ADODB.Recordset
adoM2.Fields.Append "Comercio", adInteger, 2
adoM2.Fields.Append "NoOrden", adInteger, 2
adoM2.Fields.Append "NoDepend", adInteger, 2
adoM2.Fields.Append "ValorME", adSingle
adoM2.Fields.Append "Vencim", adDate
adoM2.Fields.Append "Moned", adChar, 1
adoM2.Fields.Append "ValorP", adSingle
adoM2.Fields.Append "Pr", adBoolean
adoM2.Fields.Append "Socio", adInteger, 2
adoM2.Fields.Append "Cuota", adChar, 7
adoM2.Fields.Append "NAuto", adInteger, 2
adoM2.Fields.Append "NCobro", adInteger, 2
If nPrm = 1 Then
    adoM2.Fields.Append "NComerc", adChar, 20
ElseIf nPrm = 2 Then    'para historia(2)
    adoM2.Fields.Append "SinPagar", adSingle
End If
adoM2.CursorType = adOpenDynamic
adoM2.LockType = adLockOptimistic
adoM2.Open
End Sub




'=====================================================================
Public Sub Graba()
'=====================================================================
'Call EvitaNumOrdenDoble
adoOrdenes.AddNew

adoOrdenes!ord_nrosoc = vlNroSoc
adoOrdenes!ORD_NroCom = vlNroComerc
adoOrdenes!ord_NroOrden = vlNroOrden
adoOrdenes!ORD_DEPEND = vlNroDepend
adoOrdenes!ord_cuota = vnsCuota
adoOrdenes!ord_FEmis = vdFEmis
adoOrdenes!ord_FVto = vdFVto
adoOrdenes!ORD_PLAN = vnPlan
adoOrdenes!ord_ctasPagas = vnCtasPaga
adoOrdenes!ord_EntCta = vnsEntCta
adoOrdenes!ord_Recarg = vnsRecarg
adoOrdenes!ord_Mon = vsMoneda
adoOrdenes!ord_mecuota = vnsMECuota
adoOrdenes!ORD_MEPagos = vnsMEPagos
adoOrdenes!ord_cerro = vdCerro                  '(Cerró)
adoOrdenes!ord_tipo = 0
adoOrdenes!ORD_Func = vsFunc
adoOrdenes!ORD_FDia = vsFDia
adoOrdenes!ORD_FHora = vsFHora
adoOrdenes.Update
End Sub
        
 '=====================================================================
Public Sub GrabaEnAdoOrdns()
'=====================================================================
'Call EvitaNumOrdenDoble
On Error GoTo Problemas

mGlob.adoOrdns.AddNew

mGlob.adoOrdns!ord_nrosoc = vlNroSoc
mGlob.adoOrdns!ORD_NroCom = vlNroComerc
mGlob.adoOrdns!ord_NroOrden = vlNroOrden
mGlob.adoOrdns!ORD_DEPEND = vlNroDepend
mGlob.adoOrdns!ord_cuota = vnsCuota
mGlob.adoOrdns!ord_FEmis = vdFEmis
mGlob.adoOrdns!ord_FVto = vdFVto
mGlob.adoOrdns!ORD_PLAN = vnPlan
mGlob.adoOrdns!ord_ctasPagas = vnCtasPaga
mGlob.adoOrdns!ord_EntCta = vnsEntCta
mGlob.adoOrdns!ord_Recarg = vnsRecarg
mGlob.adoOrdns!ord_Mon = vsMoneda
mGlob.adoOrdns!ord_mecuota = vnsMECuota
mGlob.adoOrdns!ORD_MEPagos = vnsMEPagos
mGlob.adoOrdns!ord_cerro = vdCerro                  '(Cerró)
mGlob.adoOrdns!ord_tipo = 0
mGlob.adoOrdns!ORD_Func = vsFunc
mGlob.adoOrdns!ORD_FDia = vsFDia
mGlob.adoOrdns!ORD_FHora = vsFHora

mGlob.adoOrdns.Update

Exit Sub
Problemas:
MsgBox "Error 37723: Problema al grabar el registro. Error No: " & Err.Number & vbNewLine & _
    Err.Description, vbInformation, "Problemas."
fColocaMarcaAccion 40, "ErrorGrabar", "Problemas al grabar registro " & vlNroSoc, "", ""
 
End Sub
        
'NO LO ESTOY UTILIZANDO
'Private Sub EvitaNumOrdenDoble()
'Para evitar que se grabe 2 veces el mismo
'Numero de orden
'    Dim Encuentra As Boolean
'    Encuentra = True
'    Set adoB = New ADODB.Recordset
'    Set adoB.ActiveConnection = adoConn
'    adoB.Open "SELECT * FROM TBL_Ordenes ORDER BY ord_NroOrden", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
'    Do While Encuentra = True
'        adoB.MoveFirst
'        adoB.Find ("ORD_NroOrden =" & CStr(vlNroOrden))
'        If adoB.EOF Then
'            Encuentra = False
'        Else
'            Encuentra = True
'            vlNroOrden = vlNroOrden + 1
'        End If
'    Loop
'    adoB.Close
'    Set adoB = Nothing

'End Sub
  
        
        
'=====================================================================
Public Function fTomaPresupActual() As String
'=====================================================================
    Dim nMes As Byte
    Dim nAnio As Integer
    
    Set adoM2 = New ADODB.Recordset
    Set adoM2.ActiveConnection = adoConn
    adoM2.Open "SELECT * FROM TBL_Parametros", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoM2.MoveFirst
    fTomaPresupActual = adoM2!PRM_Prspst
    
    nMes = CInt(Right(fTomaPresupActual, 2))
    nAnio = CInt(Left(fTomaPresupActual, 4))
    
    vdPrsptoTermina = CDate("10/" & nMes & "/" & nAnio)
    
    adoM2.Close
    Set adoM2 = Nothing
End Function


'=====================================================================
Public Function mfSumaAdoM2(sMome As Single, sMome1 As Single) As Boolean
'=====================================================================
    'Dim sMome As Single         'suma de los vencimientos en este presupuesto
    'Dim sMome1 As Single        'suma de los vtos en otro presupuesto
    
On Error GoTo miErrr3688

    
    With adoM2
    If .RecordCount < 1 Then
        mfSumaAdoM2 = True
        Exit Function
    End If
    .MoveFirst
        Do While Not .EOF
            'es de este presupuesto
            If adoM2!pr = 0 Then
                sMome = sMome + adoM2!valorp
            Else
            'es de otro presupuesto
                sMome1 = sMome1 + adoM2!valorp
            End If
            .MoveNext
        Loop
    End With
    mfSumaAdoM2 = True
    Exit Function
miErrr3688:
MsgBox ("ERROR 3688: " & Err.Description & " ENo." & Err.Number)
mfSumaAdoM2 = False
End Function

'=====================================================================
Public Sub msOrdena1()
'=====================================================================
adoM2.Sort = "VENCIM, NOORDEN"
adoM2.MoveFirst
End Sub

'no estoy utilizando
'=====================================================================
Public Function mfTomaNumOrden() As Long
'=====================================================================
    Dim lmome As Long
    
    Set adoNumOrden = New ADODB.Recordset
    Set adoNumOrden.ActiveConnection = adoConn
    If adoNumOrden.State = adStateOpen Then adoNumOrden.Close
    adoNumOrden.Open "SELECT * FROM TBL_Parametros", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoNumOrden.MoveFirst
    lmome = adoNumOrden("NroOrden") + 1
    adoNumOrden.Close
    Set adoNumOrden = Nothing
    mfTomaNumOrden = lmome
End Function



'no estoy utilizando
'=====================================================================
Public Sub msGrabaNumOrden()
'=====================================================================
    Dim lmome As Long
    
    Set adoNumOrden = New ADODB.Recordset
    Set adoNumOrden.ActiveConnection = adoConn
    If adoNumOrden.State = adStateOpen Then adoNumOrden.Close
    adoNumOrden.Open "SELECT * FROM TBL_Parametros", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoNumOrden.MoveFirst
    lmome = adoNumOrden("NroOrden") + 1
    adoNumOrden("NroOrden") = lmome
    adoNumOrden.Update
    adoNumOrden.Close
    Set adoNumOrden = Nothing
End Sub


'=====================================================================
Public Function mfTomaUltDiaPresup() As Byte
'=====================================================================
    Dim bMome As Byte
    Set adoNumOrden = New ADODB.Recordset
    Set adoNumOrden.ActiveConnection = adoConn
    adoNumOrden.Open "SELECT * FROM TBL_Parametros", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoNumOrden.MoveFirst
    
    bMome = adoNumOrden("prm_prsphst")
    adoNumOrden.Close
    Set adoNumOrden = Nothing
    mfTomaUltDiaPresup = bMome
End Function





'FUNCION DESHABILITADA
'=====================================================================
Public Function fCalcOrdenesAPagarDelSocio() As Single
'=====================================================================
'funcion deshabilitada.
    'CALCULA EN PESOS LAS ORDENES QUE DEBE PAGAR EN EL SIG.MES
    Dim sMome As Single
    
    sMome = 0
    
      
    
    
    If adoOrdenes.RecordCount = 0 Then
        fCalcOrdenesAPagarDelSocio = 0
        Exit Function
    End If
    sPresupActual = fTomaPresupActual
    'recorre el ado
    adoOrdenes.MoveFirst
    Do While Not adoOrdenes.EOF
        sMome = sMome + fSumaOrdenes1
        adoOrdenes.MoveNext
    Loop
    fCalcOrdenesAPagarDelSocio = sMome
End Function




'FUNCION DESHABILITDA
'=====================================================================
Private Function fSumaOrdenes1() As Single
'=====================================================================
    
    Dim sMome As Single
    
   
    
    '2) VE SI vencimiento PERTENECE AL PRESUPUESTO ACTUAL
    Dim bEsPresup As Boolean
    Dim nMes As Byte
    Dim nAnio As Integer
    
    bEsPresup = False
    nMes = CInt(Right(sPresupActual, 2))
    nAnio = CInt(Left(sPresupActual, 4))
    
    vdPrsptoTermina = CDate("10/" & nMes & "/" & nAnio)
    If adoOrdenes!otro > vdPrsptoTermina Then
        bEsPresup = False
    Else
        bEsPresup = True
    End If
    
    'If Day(dVencim) > vpnPrspHst Then           'presup comienza 11/(m-1)
    '    'MES + 1 Y AÑO DEBEN COINCIDIR
    '    If Month(dVencim) + 1 = nMes And Year(dVencim) = nAnio Then
    '        bEsPresup = True
    '    End If
    'Else                                        'presup vence 10/m
    '    'MES Y AÑO DEBEN COINCIDIR
    '    If Month(dVencim) = nMes And Year(dVencim) = nAnio Then
    '        bEsPresup = True
    '    End If
    'End If
    
    '3) Toma el valor
    'toma el saldo en pesos
    If adoOrdenes!ord_Mon = "P" Then
        sMome = adoOrdenes!ord_cuota * (adoOrdenes!ORD_PLAN - adoOrdenes!ord_ctasPagas) + adoOrdenes!ord_Recarg - adoOrdenes!ord_EntCta
        'si el saldo es mmayor que la cuota, toma la cuota
        If sMome > adoOrdenes!ord_cuota Then
            sMome = adoOrdenes!ord_cuota
        End If
    Else ' en otra moneda
        sMome = adoOrdenes!ord_mecuota * (adoOrdenes!ORD_PLAN - adoOrdenes!ord_ctasPagas) + adoOrdenes!ord_Recarg - adoOrdenes!ord_EntCta
        'si el saldo es mmayor que la cuota, toma la cuota
        If sMome > adoOrdenes!ord_mecuota Then
            sMome = adoOrdenes!ord_mecuota
        End If
        'lo traduce a pesos
        cTC.msInicia
        cTC.msInicializaVariables
        cTC.vdFecha = Date      ' fecha de hoy
        cTC.msDeterminaMoneda (adoOrdenes!ord_Mon)
        sMome = sMome * cTC.mfBuscaTCambio
        cTC.msTermina
    End If
    
    '5) AGREGA EL VALOR
    If bEsPresup Then
        fSumaOrdenes1 = sMome
    Else
        fSumaOrdenes1 = 0
    End If
End Function





'=====================================================================
Public Function mfActualizaLaOrden(lNroOrden As Long, tMon As String, _
    sValorP As Single, sValorME As Single, lNroSocio As Long, _
    dVto As Date) As Byte
'=====================================================================
    
    Dim nTipoPago As Byte
    Dim sSaldo As Single

    With adoOrdenes
    'si es una cuota/ayuda
    If lNroOrden = 1 Or lNroOrden = 2 Then
        'ubica la orden
        If Not fBuscaUnaCuotaOAyuda(lNroSocio, lNroOrden, dVto) Then
            mfActualizaLaOrden = 0
            Exit Function
        End If
    
    Else    'es una orden comun
        'ubica la orden
        If Not fBuscaUnaOrden(lNroOrden) Then
            mfActualizaLaOrden = 0
            Exit Function
        End If
    End If
    'es en pesos
    If Trim(tMon) = "" Or Trim(tMon) = "P" Then
        'paga una cuota
        If sValorP = !ord_cuota Then
            !ord_ctasPagas = !ord_ctasPagas + 1
            nTipoPago = 5
        'es una entrega a cuenta
        Else
            !ord_EntCta = !ord_EntCta + sValorP
            nTipoPago = 6
        End If
    
    'es en moneda extranjera
    Else
            'paga una cuota en dolares
        If sValorME = !ord_mecuota Then
            !ord_ctasPagas = !ord_ctasPagas + 1
            nTipoPago = 5
        'es una entrega a cuenta en dolares
        Else
            !ord_EntCta = !ord_EntCta + sValorME
            nTipoPago = 6
        End If
    
    End If
    .Update
    
    'cierra la orden
    If Trim(tMon) = "" Or Trim(tMon) = "P" Then
        sSaldo = !ord_cuota * (!ORD_PLAN - !ord_ctasPagas) + !ord_Recarg - !ord_EntCta
    Else
        sSaldo = !ord_mecuota * (!ORD_PLAN - !ord_ctasPagas) + !ord_Recarg - !ord_EntCta
    
    End If
    If sSaldo < 1 And sSaldo > -1 Then
        !ord_cerro = Format(Date, "short date")
        .Update
    End If
    End With
    mfActualizaLaOrden = nTipoPago
End Function
Public Function mfAgregaRecargoEnOrdenes(sPrm As Single) As Boolean
adoOrdenes!ord_Recarg = adoOrdenes!ord_Recarg + sPrm
adoOrdenes.Update
mfAgregaRecargoEnOrdenes = True
End Function

'=====================================================================
Public Sub msPreparaOrdenesEnAdoM2()
'=====================================================================
    'OJO: LA ORDEN POR EL VALOR TOTAL al sacar la orden
    'es para la Historia del socio
    'Sin recargo ni entregas a cuentas.
    
    '2) Crea la tabla virtual en el adoM2
    Call mfMOAPCreaTablaVirtual(1)
    
    '3) recorre el ado
    If adoOrdenes.RecordCount < 1 Then
        MsgBox "Sin Registros"
        Exit Sub
    End If
    With adoOrdenes
        .MoveFirst
        Do While Not .EOF
                Call mfLlenaUnRegAdoM2
            
            .MoveNext
        Loop
    End With
 End Sub

'=====================================================================
Private Sub mfLlenaUnRegAdoM2()
'=====================================================================
    'OO: LA ORDEN POR EL TOTAL
    adoM2.AddNew
    With adoOrdenes

        adoM2!comercio = !ORD_NroCom
        adoM2!NoOrden = !ord_NroOrden
        adoM2!NoDepend = !ORD_DEPEND
        adoM2!vencim = !ord_FEmis
        adoM2!socio = !ord_nrosoc
        adoM2!nAuto = !ord_Auto
        If !ord_Mon = "P" Then
            adoM2!valorp = !ord_cuota * !ORD_PLAN   '+ !ord_Recarg - !ord_EntCta
            adoM2!valorME = 0
        Else
            adoM2!valorp = 0
            adoM2!valorME = !ord_mecuota * !ORD_PLAN    '+ !ord_Recarg - !ord_EntCta
        End If
        adoM2!moned = !ord_Mon
        'adoM2!valorp = mValorME * cTC.mfDevuelveCambio(!ord_Mon, Date)
        adoM2!ncomerc = ""
        adoM2!cuota = !ORD_PLAN
    End With
    adoM2.Update
End Sub



'=====================================================================
Public Function fBusca2TodasLasOrdenesUnSocio() As Boolean
'=====================================================================
  'para historia(2)
    
    On Error GoTo merr5247
    
    Dim sMom As String
    'Dim dMM As Date
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT *,cdate('01/01/1900')  as Otro FROM TBL_Ordenes " & _
        "WHERE ORD_NroSoc =" & _
        CStr(vlNroSoc) & ";"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount < 1 Then
        fBusca2TodasLasOrdenesUnSocio = False
        Exit Function
    End If
    Dim nMM As Integer
    With adoOrdenes
    .MoveFirst
        
        Do While Not .EOF
            'coloca la fecha de vencimiento en el campo otro
            'Debug.Print !ord_NroOrden & "  " & !otro
            !otro = !ord_FVto
            .MoveNext
        Loop
    End With

   fBusca2TodasLasOrdenesUnSocio = True
    Exit Function
    
merr5247:
    MsgBox "ERROR 5247: " & Err.Description & " EN: " & Err.Number
    fBusca2TodasLasOrdenesUnSocio = False
End Function

'=====================================================================
Public Sub ms2PreparaOrdenesAPagarEnAdoM2()
'=====================================================================
    'para historia(2)
    '1)CALCULA EN PESOS LAS ORDENES QUE DEBE PAGAR EN EL SIG.MES
    '2)LAS PONE EN UNA TABLA

   
    
    '2) Crea la tabla virtual en el adoM2
    Call mfMOAPCreaTablaVirtual(2)
    
    '3) recorre el ado
    Dim nM1 As Single       'lo que pagó
    
    With adoOrdenes
        .MoveFirst
        Do While Not .EOF
                'EL TOTAL PAGADO
                If !ord_Mon = "P" Then
                    nM1 = !ord_ctasPagas * !ord_cuota + !ord_EntCta - !ord_Recarg
                Else
                    nM1 = !ord_ctasPagas * !ord_mecuota + !ord_EntCta - !ord_Recarg
                End If
            'RECORRE LAS CUOTAS
            Dim nNCuot As Integer
            Dim nPlan As Integer     'por el total de cuotas
            nNCuot = !ORD_PLAN
            nPlan = nNCuot
            Do While nNCuot > 0
                '3.1) Realiza los calculos=
                '3.2) Los guarda en la tabla
                If nM1 < 0 Then nM1 = 0
                Call mfPorCuota(nNCuot, nM1)
                If !ord_Mon = "P" Then
                    nM1 = nM1 - !ord_cuota
                Else
                    nM1 = nM1 - !ord_mecuota
                End If
                '3.3) Realiza la suma
                'sMome = sMome + fSumaOrdenes1
                nNCuot = nNCuot - 1
            Loop
            .MoveNext
        Loop
    End With
    
  '4) Ordena la tabla
  Me.msOrdena1
  
 
End Sub
'=====================================================================
Private Sub mfPorCuota(bNroCuota As Integer, nS As Single)
'=====================================================================
    'para el ms2PreparaOrdenesAPagarEnAdoM2
    Dim dMM As Date
    
    '1)
    adoM2.AddNew
    adoM2!comercio = adoOrdenes!ORD_NroCom
    adoM2!NoOrden = adoOrdenes!ord_NroOrden
    adoM2!NoDepend = adoOrdenes!ORD_DEPEND
    adoM2!vencim = adoOrdenes!otro
    adoM2!socio = adoOrdenes!ord_nrosoc
    adoM2!nAuto = adoOrdenes!ord_Auto
    
    
    With adoOrdenes
            
            '2)Cuota
            
            adoM2!valorp = !ord_cuota
            adoM2!cuota = CStr(!ORD_PLAN - bNroCuota + 1) & "/" & CStr(!ORD_PLAN)
            
            '3) Es en moneda extranjera
            If Not !ord_Mon = "P" Then
                If nS < !ord_mecuota Then
                    adoM2!SinPagar = !ord_mecuota - nS
                Else
                    adoM2!SinPagar = 0
                End If
                adoM2!valorp = !ord_mecuota
                adoM2!moned = !ord_Mon
                'adoM2!valorp = !ord_CuotaME * cTC.mfDevuelveCambio(!ord_Mon, Date)
            Else        'en pesos
                If nS < !ord_cuota Then
                    adoM2!SinPagar = !ord_cuota - nS
                Else
                    adoM2!SinPagar = 0
                End If

                adoM2!moned = ""
            End If
            
 
    
    '5)Cambia el vencimiento (para el proximo registro)
    If Month(!otro) = 12 Then
        'dd/1/aa+1
        dMM = CDate(Format(Day(!otro) & "/1/" & Year(!otro) + 1, "short date"))
    Else
        'dd/mm+1/aa
        dMM = CDate(Format(Day(!otro) & "/" & Month(!otro) + 1 & "/" & Year(!otro), "short date"))
    End If
    !otro = dMM         'lo cambia en el propio adoOrdenes
    
    '6)
    adoM2.Update
   End With

End Sub
'=====================================================================
Public Function fBuscaOrdenesTodosSocios5() As Boolean
'=====================================================================
    'Todas las ordenes:
    'Incluso las cerradas
    'NO la anuladas
    
    On Error GoTo merr5235
    
    Dim sMom As String
    Dim sSaldo As Single
  

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT *, cdate('01/02/1900') as Otro " & _
        "FROM TBL_Ordenes " & _
        "WHERE NOT ord_tipo = 4;"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount < 1 Then
        MsgBox "Sin Ordenes"
        fBuscaOrdenesTodosSocios5 = False
        Exit Function
    End If
      Call ms5ColocaFechaVtoEnOtro

    fBuscaOrdenesTodosSocios5 = True
    Exit Function
    
merr5235:
    MsgBox "ERROR 5235: " & Err.Description & " EN: " & Err.Number
    fBuscaOrdenesTodosSocios5 = False
End Function

'=====================================================================
Public Function fBuscaOrdenesTodosSocios() As Boolean
'=====================================================================
    'On Error GoTo merr5233
    
    Dim sMom As String
    Dim sSaldo As Single
  

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT *, cdate('01/02/1900') as Otro " & _
        "FROM TBL_Ordenes " & _
        "WHERE year(ord_cerro) < 1901;"
    adoOrdenes.Open sMom, adoConn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount < 1 Then
        fBuscaOrdenesTodosSocios = False
        Exit Function
    End If
    Call msColocaFechaVtoEnOtro

    fBuscaOrdenesTodosSocios = True
    Exit Function
    
merr5233:
    MsgBox "ERROR 5233: " & Err.Description & " EN: " & Err.Number
    fBuscaOrdenesTodosSocios = False
End Function
'=====================================================================
Public Sub ms8PreparaOrdenesAPagarEnAdoM2(nPrm As Byte)
'=====================================================================
    '1)CALCULA EN PESOS LAS ORDENES QUE DEBE PAGAR EN EL SIG.MES
    '2)LAS PONE EN UNA TABLA
    Dim vvNoCuota As Byte

   
    
    '1)Toma el mes del Presupuesto Actual
    Dim nMes As Byte
    Dim nAnio As Integer
    'OJO abre y cierra el adoM2
    sPresupActual = fTomaPresupActual
   
    nMes = CInt(Right(sPresupActual, 2))
    nAnio = CInt(Left(sPresupActual, 4))
    
    vdPrsptoTermina = CDate("10/" & nMes & "/" & nAnio)
    
    '2) Crea la tabla virtual en el adoM2
    Call mfMOAPCreaTablaVirtual(nPrm)
    
    '3) recorre el ado
    Dim nM1 As Single
    Dim nCoc As Integer
    Dim nRes As Single
    
    With adoOrdenes
        .MoveFirst
        Do While Not .EOF
            'Para el nPrm=5 no coloca ordenes vencidas
                    'Numero de pagos
                    'Sin recargo o entCta
                    If !ord_Recarg = !ord_EntCta Then
                        nCoc = !ORD_PLAN - !ord_ctasPagas
                    'hay recargo o entcTA
                    Else
                        'EN PESOS
                        If !ord_Mon = "P" Then
                            nM1 = (!ORD_PLAN - !ord_ctasPagas) * !ord_cuota + !ord_Recarg - !ord_EntCta
                            'por si hubo entregas o recargos
                             nCoc = Fix(nM1 / !ord_cuota)    'PARTE ENTERA
                            nRes = nM1 - nCoc * !ord_cuota     'RESTO
                            If !ord_Recarg < !ord_EntCta Then
                                If Abs(nRes) > 10 Then nCoc = nCoc + 1
                            End If
                        Else
                            nM1 = (!ORD_PLAN - !ord_ctasPagas) * !ord_mecuota + !ord_Recarg - !ord_EntCta
                            'por si hubo entregas o recargos
                            nCoc = Fix(nM1 / !ord_mecuota)
                            nRes = nM1 - nCoc * !ord_mecuota     'RESTO
                            If !ord_Recarg < !ord_EntCta Then
                                If Abs(nRes) > 2 Then nCoc = nCoc + 1
                            End If
                        End If
                        If !ord_Recarg < !ord_EntCta Then
                            If Not nRes = 0 Then nCoc = nCoc + 1
                        End If
                    End If
                    'RECORRE LAS CUOTAS
                    Dim nTot As Integer     'por el total de cuotas
                    nTot = nCoc
                    Do While nCoc > 0
                        '3.1) Realiza los calculos=
                        '3.2) Los guarda en la tabla
                        Call mf8MOAP31(nCoc, nTot, nPrm)
                        '3.3) Realiza la suma
                        'sMome = sMome + fSumaOrdenes1
                        nCoc = nCoc - 1
                    Loop
                    .MoveNext
        Loop
    End With
    
  '4) Ordena la tabla
  Me.msOrdena1
  
 
End Sub
'=====================================================================
Private Sub mf8MOAP31(bNroCuota As Integer, nTot As Integer, nPrm As Byte)
'=====================================================================
    Dim mValor As Single
    Dim mValorME As Single
    Dim dMM As Date
    
    '1)
    adoM2.AddNew
    adoM2!comercio = adoOrdenes!ORD_NroCom
    adoM2!NoOrden = adoOrdenes!ord_NroOrden
    'adoM2!NoDepend = adoOrdenes!ORD_DEPEND
    adoM2!vencim = adoOrdenes!otro
    adoM2!socio = adoOrdenes!ord_nrosoc
    adoM2!nAuto = adoOrdenes!ord_Auto
    
    With adoOrdenes
    '2)Cuota
    If bNroCuota = nTot Then       'ES LA PRIMERA
        mValor = !ord_Recarg - !ord_EntCta
        If mValor = 0 Then
            mValor = !ord_cuota
            mValorME = !ord_mecuota
        Else
            mValor = (!ORD_PLAN - !ord_ctasPagas - nTot + 1) * !ord_cuota + !ord_Recarg - !ord_EntCta
            mValorME = (!ORD_PLAN - !ord_ctasPagas - nTot + 1) * !ord_mecuota + !ord_Recarg - !ord_EntCta
        End If
    Else                        'NO ES LA PRIMERA
        mValor = !ord_cuota
        mValorME = !ord_mecuota
    End If
    'adoM2!valorp = mValor
    'adoM2!cuota = CStr(!ORD_PLAN - bNroCuota + 1) & "/" & CStr(!ORD_PLAN)
    
    'es en pesos
    If nPrm = 1 Then
        adoM2!valorp = mValor
    Else
        adoM2!valorp = mValorME
    End If
    '3) Es en moneda extranjera
    'If Not !ord_Mon = "P" Then
     '   adoM2!valorME = mValorME
     '   adoM2!moned = !ord_Mon
     '   adoM2!valorp = mValorME * cTC.mfDevuelveCambio(!ord_Mon, Date)
    'End If
    
    '4)Es de este presupuesto
     If adoOrdenes!otro > vdPrsptoTermina Then
        adoM2!pr = True
    Else
        adoM2!pr = False
     End If
    
    '5)Cambia el vencimiento (para el proximo registro)
    'MsgBox Day(!otro)
    'MsgBox Month(!otro)
    'MsgBox Year(!otro)
    If Month(!otro) = 12 Then
        dMM = CDate(Format(Day(!otro) & "/1/" & Year(!otro) + 1, "short date"))
        Else
        dMM = CDate(Format(Day(!otro) & "/" & Month(!otro) + 1 & "/" & Year(!otro), "short date"))
    End If
    !otro = dMM         'lo cambia en el propio adoOrdenes
    'If nprm = 1 Then
    '    adoM2!ncomerc = "" & Left(Trim(!NombCom), 15)
    'End If
    End With
    
    '6)
    adoM2.Update
End Sub
Public Function TomaYGrabaNumOrden() As Long
    Dim lmome As Long
    On Error GoTo Mal
    
        
    Set adoNumOrden = New ADODB.Recordset
    'If adoNumOrden.State = adStateOpen Then adoNumOrden.Close
    adoNumOrden.Open "SELECT * FROM TBL_Parametros", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
        
    adoNumOrden.MoveFirst
    lmome = adoNumOrden("NroOrden") + 1
    adoNumOrden("NroOrden") = lmome
    adoNumOrden.Update
    DoEvents
    
    
    'busca que no este repetido
    'Dim Encuentra As Boolean
    'Encuentra = True
    
    'Do While Encuentra = True
    '    adoB.MoveFirst
    '    adoB.Find ("ORD_NroOrden =" & CStr(cOrd.vlNroOrden))
    '    If adoB.EOF Then
    '        Encuentra = False
    '    Else
    '        Encuentra = True
    '        cOrd.vlNroOrden = cOrd.vlNroOrden + 1
    '    End If
    'Loop
    'If Not cOrd.vlNroOrden = lmome Then    'cambio
    '    adoNumOrden("NroOrden") = cOrd.vlNroOrden
    '    lmome = cOrd.vlNroOrden
    '    adoNumOrden.Update
    'End If
    'DoEvents
    TomaYGrabaNumOrden = lmome
    If adoNumOrden.State = adStateOpen Then adoNumOrden.Close
    Set adoNumOrden = Nothing
    Exit Function
Mal:
    MsgBox "Error 48236: al tomar No de Orden. No: " & Err.Number & vbNewLine & _
        Err.Description, vbExclamation, "Problemas"
    fColocaMarcaAccion 40, "Toma No Orden", "Problemas al tomar No. Orde", cOrd.vlNroSoc, ""
    TomaYGrabaNumOrden = 0
   If adoNumOrden.State = adStateOpen Then adoNumOrden.Close
    Set adoNumOrden = Nothing
End Function
