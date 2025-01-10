Attribute VB_Name = "mGlob"
Option Explicit

Global adoConn As New ADODB.Connection
Dim adoI As New ADODB.Command           'jv
Global adoM As New ADODB.Recordset      ' Auxliar (solo lectura)
Global adoOrdns As New ADODB.Recordset     ' La tabla ORDENES: es muy grande, la voy a abrir solo una vez
Global cOrd As New clsOrdenes                   ' Idem
Global cPgs As New clsPagos


Global vpMuestraTabla As Byte           'jv
Global vpbVieneDeMuestraTabla As Boolean
Global vplNumeroRegistro As Variant
Global Const kMstrSocAlf = 1         'jv
Global Const kMstrSoc1 = 2        'jv
Global Const kMuestraComercios = 3
Global Const kMstrSocAlf3 = 4
Global Const kMstrSocPorCobr1 = 5
Global Const kMstrSocAlf5 = 7
Global Const kMuestraDepend = 6
Global Const kMstrSocAlf6 = 8    'fjInformes: Informe de un socio
Global Const kMstrSocAlf9 = 9
Global Const kMstrSocPorNC2 = 10
Global Const kMstrSocAlf7 = 11
Global Const kMstrSocPorNC3 = 12
Global Const kMstrComerc2 = 13
Global Const kMstrComerc3 = 14
Global Const kMstrGastos = 15
Global Const kMstrComerc4 = 16

Global Const kNivel1 = 1
Global Const kNivel2 = 2
Global Const kNivel3 = 3
Global Const kNivel4 = 4
Global Const kNivel5 = 5
Global Const kNivel6 = 6


Global Const kMenorAñoTrabajo = 1995
Global Const kMayorAñoTrabajo = 2015


'LA FUNCION QUE SE ESTA REALIZANDO O
'SE VA A REALIZAR
'Varios movimientos en un mismo formulario...........
Global vpFormMovim As Byte
Global Const kFormIngresa = 1
Global Const kFormModifica = 2
Global Const kFormElimina = 3
Global Const kFormMuestra = 4
Global Const kFormAnula = 5
Global Const kFormRecorre = 6
Global Const kFormMira = 7
Global vpbCancel As Boolean
Global vpsPideDato As String

'FUNCIONARIO ........................................
Global vpnFuncionario As Byte         'numero funcionario
Global vpnNivelFuncionario As Byte
Global vptNombreFuncionario As String
Global vptFuncPass As String
Global vptMesPresup As String   'mes de presupuesto aaaamm
Global vpnPrspHst As Byte      'dia que vence el presupuesto
Global vplNroSocio As Long
Global vplNroOrden As Long
Global vplNroRecibo As Long
Global vpsCuotaSAct As Single
Global vpsCuotaSHon As Single
Global vpsCuotaSCop As Single
Global vpsAyuda As Single
Global vpsRecargo As Single         '???????
Global vpnAñoOperac As Integer
Global vpbMesOperac As Byte
Global vpsPorcRecarg As Single      'porcentaje recargo
Global vptReporte As String         'nombre archivo para mireporte


Global Const kYO = 45

'Moneda extranjera ..................................
Global vpTCD As Single          't/c dolar
Global vpTCR As Single          't/c Reales
Global vpTCA As Single          't/c pesos arg
Global vpTCU As Single          't/c Unid reaj


Public Function mensaje()
MsgBox "El registro se ha ingresado satisfactoriamente", vbExclamation, "Circulo Policial"
End Function


Public Function msgtablas()
MsgBox "Debe ingresar un registro", vbError, "Circulo Policial"
End Function


Public Sub HistoriaEntra()      'jv
    Set adoI.ActiveConnection = adoConn
    'adoI.CommandText = "insert into TBL_Historia values('" & vpnFuncionario & "','" & Date & "','" & Time & ",'A')"
    adoI.CommandText = "insert into TBL_Historia values('" & vpnFuncionario & "','" & Date & "','" & Time & "','E')"
    adoI.Execute
End Sub


Public Sub HistoriaSale()   'jv
    Set adoI.ActiveConnection = adoConn
    adoI.CommandText = "insert into TBL_Historia values('" & vpnFuncionario & "','" & Date & "','" & Time & "','S')"
    adoI.Execute
End Sub

Public Sub msCargaComboSituacLaboral(cPrm As ComboBox)
'CARGA LA TABLA SITUACION LABORAL (SLABORAL)
'EN UN COMBO BOX
Dim i As Byte
   Set adoM = New ADODB.Recordset
    adoM.Open "select * from slaboral", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoM.RecordCount <> 0 Then
        For i = 1 To adoM.RecordCount
            cPrm.AddItem (adoM!Desc)
            adoM.MoveNext
        Next i
    End If
    adoM.Close
    Set adoM = Nothing
End Sub



Public Function mfDevCodSitLaboral(sprm As String) As Byte

'Buscar código Situac.Laboral
    Set adoM = New ADODB.Recordset
    adoM.Open "select * from slaboral", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoM.RecordCount <> 0 Then
        adoM.MoveFirst
        'adoM.Find (Trim(UCase(adoM!Desc)) = Trim(UCase(sPrm)))
        adoM.Find ("Desc ='" & sprm & "'")
        If Not adoM.EOF Then
            mfDevCodSitLaboral = adoM!idsitlab
        End If
    Else
        mfDevCodSitLaboral = 0
    End If
    adoM.Close
    Set adoM = Nothing
End Function







'======================================================
Public Function TomaDatosFUncionario(sprm As String) As Boolean
'======================================================
On Error GoTo jj001
    Dim sCriterio As String
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM TBL_Funcio", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    sCriterio = "CodSeg =" & CInt(sprm)
    adoM.MoveFirst
    adoM.Find (sCriterio)
    'si no la encuentra
    If adoM.EOF Or adoM.BOF Then
        TomaDatosFUncionario = False
    ElseIf IsNull(adoM!nombre) Then
        TomaDatosFUncionario = False
    Else
        vpnFuncionario = CInt(sprm)
        vptFuncPass = adoM!clave
        vpnNivelFuncionario = adoM!nivel
        vptNombreFuncionario = adoM!nombre
        MDIingreso.Caption = "Círculo Policial.            Func: " & adoM!nombre
        TomaDatosFUncionario = True
    End If
    adoM.Close
    Set adoM = Nothing
Exit Function
jj001:
MsgBox ("ERROR jj001: " & Err.Description)
End
End Function

'======================================================
Public Function CambioDeContraseña(sprm As String) As Boolean
'======================================================
On Error GoTo jj002
    Dim sCriterio As String
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM TBL_Funcio", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    sCriterio = "CodSeg =" & vpnFuncionario
    adoM.MoveFirst
    adoM.Find (sCriterio)
    'si no la encuentra
    If adoM.EOF Or adoM.BOF Then
        CambioDeContraseña = False
    Else
        adoM!clave = sprm
        adoM.Update
        vptFuncPass = sprm
        CambioDeContraseña = True
    End If
    adoM.Close
    Set adoM = Nothing
Exit Function
jj002:
MsgBox ("ERROR jj002: " & Err.Description)
CambioDeContraseña = False
End
End Function
'=====================================================
Public Function fTomaMesOperacYParametros() As Boolean
'=====================================================
On Error GoTo ja001
    adoM.Open "SELECT * FROM TBL_Parametros", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoM.MoveFirst
    
    vpnAñoOperac = CInt(Left(adoM("prm_prspst"), 4))
    vpbMesOperac = CByte(Right(adoM("prm_prspst"), 2))
    vptMesPresup = adoM("prm_prspst")   'aaaamm
    vpnPrspHst = CByte(adoM("PRM_PRSPHST")) 'dd
    vplNroOrden = CLng(adoM("NroOrden"))
    vplNroRecibo = CLng(adoM("prm_NroRecibo"))
    vpsCuotaSAct = CSng(adoM("prm_SocAct"))
    vpsCuotaSHon = CSng(adoM("prm_SocHon"))
    vpsCuotaSCop = CSng(adoM("prm_SocCop"))
    vpsPorcRecarg = CSng(adoM("prm_Recarg"))
    vpsAyuda = CSng(adoM("prm_Ayuda"))
    vpsRecargo = CSng(adoM("prm_recarg"))
    adoM.Close
    Set adoM = Nothing
    fTomaMesOperacYParametros = True
Exit Function
ja001:
MsgBox ("ERROR ja001: " & Err.Description)
fTomaMesOperacYParametros = False
End
End Function


Public Sub msGuardaFalla(lSocio As Long, lOrden As Long, _
    dFecha As Date, tDetalle As String)
    'Dim adoF As New ADODB.Recordset
    'adoF.Open "SELECT * FROM TBL_Fallas", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    'adoF.AddNew
    'adoF!fl_NroSocio = lSocio
    'adoF!fl_NroOrden = lOrden
    'adoF!fl_fecha = dFecha
    'adoF!fl_deta = tDetalle
    'adoF.Update
    'adoF.Close
    'Set adoF = Nothing
End Sub

'==================================================
Public Function rCierraMes_VerifMesEsteCompleto(nPrm As Byte) As Boolean
'==================================================
'verifica que ya no se haya realizado el cierre de mes
'nPrm=1 Cierre de mes
'nprm=2 Listado Jefatura + Jefatura.dbf

On Error GoTo ja003
    Dim miMsg As String
    
    'BUSCA EN LA TABLA TBL_ACCION
    adoM.Open "SELECT * FROM TBL_accion", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
        
    Select Case nPrm
        'QUE ESTEN LOS SIGUIENTES ACCION=No   NROIDEN=aaaamm anterior
        Case 1
            miMsg = "El Presupuesto ya fue cerrado!!!"
            If rCierraMes_VerifMesAux(21, vptMesPresup) Then GoTo saleMal
        Case 2
            miMsg = "El envío a Jefatura ya fue generado!!!"
            If rCierraMes_VerifMesAux(22, vptMesPresup) Then GoTo saleMal
        Case 3
            miMsg = "El envío a Centro Policial ya fue generado!!!"
            If rCierraMes_VerifMesAux(23, vptMesPresup) Then GoTo saleMal
        Case 6
            miMsg = "El cierre de Comercios ya fue generado!!!"
            If rCierraMes_VerifMesAux(26, vptMesPresup) Then GoTo saleMal
       Case 31
            miMsg = "El pago de Jefatura ya fue generado!!!"
            If rCierraMes_VerifMesAux(31, vptMesPresup) Then GoTo saleMal
       Case 32
            miMsg = "El pago de Círculo P. ya fue generado!!!"
            If rCierraMes_VerifMesAux(32, vptMesPresup) Then GoTo saleMal
    End Select
    
    adoM.Close
    Set adoM = Nothing
    rCierraMes_VerifMesEsteCompleto = True

Exit Function

saleMal:
adoM.Close
Set adoM = Nothing
MsgBox miMsg, , "Atención"
rCierraMes_VerifMesEsteCompleto = False
Exit Function

ja003:
adoM.Close
Set adoM = Nothing
MsgBox ("ERROR ja003: " & Err.Description)
rCierraMes_VerifMesEsteCompleto = False
End Function



'==================================================
Public Function rCierraMes_VerifMesAnteriorEsteCompleto(nPrm As Byte) As Boolean
'==================================================
On Error GoTo ja002
    Dim sEjercAnt As String

    If vpbMesOperac = 1 Then
        sEjercAnt = CStr(vpnAñoOperac - 1) & "12"
    Else
        If vpbMesOperac - 1 < 10 Then
            'sjercant= aaaa=0m
            sEjercAnt = vpnAñoOperac & "0" & vpbMesOperac - 1
        Else
            'sejercant = aaaamm
            sEjercAnt = vpnAñoOperac & vpbMesOperac - 1
        End If
    End If
    
    'BUSCA EN LA TABLA TBL_ACCION
    adoM.Open "SELECT * FROM TBL_accion", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
        
    Select Case nPrm
        Case 1
            'QUE ESTEN LOS SIGUIENTES ACCION=No   NROIDEN=aaaamm anterior
            If Not rCierraMes_VerifMesAux(21, sEjercAnt) Then GoTo saleMal
            If Not rCierraMes_VerifMesAux(22, sEjercAnt) Then GoTo saleMal
            If Not rCierraMes_VerifMesAux(23, sEjercAnt) Then GoTo saleMal
            If Not rCierraMes_VerifMesAux(24, sEjercAnt) Then GoTo saleMal
            If Not rCierraMes_VerifMesAux(25, sEjercAnt) Then GoTo saleMal
        Case 6
            'QUE ESTEN LOS SIGUIENTES ACCION=No   NROIDEN=aaaamm anterior
            If Not rCierraMes_VerifMesAux(26, sEjercAnt) Then GoTo saleMal
    End Select
    adoM.Close
    Set adoM = Nothing
    rCierraMes_VerifMesAnteriorEsteCompleto = True
Exit Function


ja002:
MsgBox ("ERROR ja002: " & Err.Description)

saleMal:
    adoM.Close
    Set adoM = Nothing
    rCierraMes_VerifMesAnteriorEsteCompleto = False

End Function



'==================================================
Private Function rCierraMes_VerifMesAux(sprm As String, sPrm2 As String) As Boolean
'==================================================
adoM.MoveFirst
adoM.Find ("acc_NroIdent ='" & sPrm2 & "'")
If adoM.EOF Then
    rCierraMes_VerifMesAux = False
    Exit Function
End If
adoM.Find ("acc_accion =" & sprm)
If adoM.EOF Then
    rCierraMes_VerifMesAux = False
    Exit Function
End If
If adoM("acc_NroIdent") = sPrm2 And _
    adoM("acc_accion") = sprm Then
    rCierraMes_VerifMesAux = True
Else
    rCierraMes_VerifMesAux = False
End If
End Function
'==================================================
Public Function fColocaMarcaAccion(nPrm As Integer, s1Prm As String, s2Prm As String, s3Prm As String, s4Prm As String) As Boolean
'==================================================
' 4 anula orden
' 8 anula cobro
'21  Prepago
'22      Disq Jef
'23      Disq CP
'24      pago Jef
'25      pago CP
' 26 cierre comercio
' 30 cambio parametros
' 40 errores
On Error GoTo ja0ab
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM TBL_accion", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
        
    adoM.AddNew
    adoM("acc_accion") = nPrm
    adoM("acc_NroIdent") = s1Prm
    adoM("acc_Detalle") = s2Prm
    adoM("acc_Text1") = s3Prm
    adoM("acc_Text2") = s4Prm
    adoM("acc_Func") = vpnFuncionario
    adoM("acc_FDia") = Format(Date, "short date")
    adoM("acc_FHora") = Format(Time, "short time")
    adoM.Update
    adoM.Close
    Set adoM = Nothing
    fColocaMarcaAccion = True
Exit Function


ja0ab:
MsgBox ("ERROR ja0ab: " & Err.Description & " Nro" & Err.Number)
fColocaMarcaAccion = False
End Function


Public Function mfEsFecha(sprm As String) As Boolean
Dim dPrm As Date
If Not IsDate(sprm) Then
    GoTo Mal
End If
dPrm = CDate(sprm)
If Year(dPrm) > 2100 Then
    GoTo Mal
End If
If Year(dPrm) < 1900 Then
    GoTo Mal
End If
GoTo bien
Mal:
mfEsFecha = False
Exit Function
bien:
mfEsFecha = True
End Function
Public Sub MouseOn()
Screen.MousePointer = vbNormal
End Sub
Public Sub MouseOff()
Screen.MousePointer = vbHourglass
End Sub

