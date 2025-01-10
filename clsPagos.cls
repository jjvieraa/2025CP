VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public adoPagos As New ADODB.Recordset

Public Function mfAbrePagos() As Boolean
    On Error GoTo mErrPagos1
    Set adoPagos = New ADODB.Recordset
    Set adoPagos.ActiveConnection = adoconn
    
    If adoPagos.State = adStateOpen Then adoPagos.Close
   
    adoPagos.Open "SELECT * FROM tbl_Pagos;", adoconn, adOpenDynamic, adLockOptimistic
    mfAbrePagos = True
    Exit Function
mErrPagos1:
    MsgBox "Error MM34: " & Err.Description & " " & Err.Number
    mfAbrePagos = False
End Function



Public Function mfGuardaUnPago(lNS As Long, _
    lNO As Long, sV As Single, dF As Date, lNR As Long, _
    lNC As Long, ByVal tM As String, sME As Single, _
    nM As Integer, ByVal tD As String, dT As Date, tF As String) As Boolean
    
    On Error GoTo mErrPagos2
    
    If Len(tD) > 30 Then tD = Left(tD, 30)
    If Trim(tM) = "" Then tM = "P"
      
    adoPagos.AddNew
    adoPagos!pag_NroSoc = lNS
    adoPagos!pag_NroOrden = lNO
    adoPagos!pag_Valor = sV
    adoPagos!pag_Fecha = dF
    adoPagos!pag_NroPago = lNR
    adoPagos!pag_NroCom = lNC
    adoPagos!pag_mon = "" & tM
    adoPagos!pag_ValME = sME
    adoPagos!pag_Motivo = nM
    adoPagos!pag_det = tD
    adoPagos!pag_Hora = Format(dT, "short time")
    adoPagos!pag_Func = tF
    adoPagos.Update
    mfGuardaUnPago = True
    Exit Function
mErrPagos2:
    MsgBox "Error MM35: " & Err.Description & " " & Err.Number
    mfGuardaUnPago = False
End Function



Public Function mfCierraPagos() As Boolean
    On Error GoTo mErrPagos3
    If adoPagos.State = adStateOpen Then adoPagos.Close
    Set adoPagos = Nothing
    mfCierraPagos = True
    Exit Function
mErrPagos3:
    MsgBox "Error MM36: " & Err.Description & " " & Err.Number
    mfCierraPagos = False
End Function

Public Function mfAbrePagosDeUnSocio(sprm As String) As Boolean
    On Error GoTo mErrPagos4
    Set adoPagos = New ADODB.Recordset
    Set adoPagos.ActiveConnection = adoconn
    
    If adoPagos.State = adStateOpen Then adoPagos.Close
   'OJO: VIENEN TAMBIEN LOS pag-motivo = 7 es lo que se genero de recargos y no el pago
    adoPagos.Open "SELECT * FROM tbl_Pagos WHERE" & _
                " pag_NroSoc =" & sprm & ";", _
                adoconn, adOpenDynamic, adLockOptimistic
    mfAbrePagosDeUnSocio = True
    Exit Function
mErrPagos4:
    MsgBox "Error MM37: " & Err.Description & " " & Err.Number
    mfAbrePagosDeUnSocio = False
End Function
