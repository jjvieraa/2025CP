VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrePag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public adoPrePag As New ADODB.Recordset

Public Function mfAbrePrePagos() As Boolean
    On Error GoTo mErrPrePagos1
    Set adoPrePag = New ADODB.Recordset
    Set adoPrePag.ActiveConnection = adoconn
    
    If adoPrePag.State = adStateOpen Then adoPrePag.Close
   
    adoPrePag.Open "SELECT * FROM tbl_PrePago;", adoconn, adOpenDynamic, adLockOptimistic
    mfAbrePrePagos = True
    Exit Function
mErrPrePagos1:
    MsgBox "Error MN34: " & Err.Description & " " & Err.Number
    mfAbrePrePagos = False
End Function



Public Function mfGuardaUnPrePago(lNS As Long, _
    lNO As Long, sV As Single, dF As Date, dV As Date, _
     bTip As Byte, sPresup As String, lAuto As Long, _
    lNC As Long, tM As String, sME As Single, _
      sT As String, tF As String) As Boolean
    
    'On Error GoTo mErrPrePagos2
    
    'If Len(tD) > 30 Then tD = Left(tD, 30)
        
    If Trim(tM) = "" Then tM = "P"
    adoPrePag.AddNew
    
    adoPrePag!pp_NroSoc = lNS
    adoPrePag!pp_NroCom = lNC
    adoPrePag!pp_NroOrden = lNO
    adoPrePag!pp_Valor = sV
    adoPrePag!pp_Femis = dF
    adoPrePag!pp_FVto = dV
    adoPrePag!pp_Mon = tM
    adoPrePag!pp_ValorME = sME
    adoPrePag!pp_Tipo = bTip
    adoPrePag!pp_Presup = sPresup
    adoPrePag!pp_Func = tF
    adoPrePag!pp_FHora = sT
    'adoPrePag!pp_Auto = lA
    adoPrePag.Update
    mfGuardaUnPrePago = True
    Exit Function
mErrPrePagos2:
    MsgBox "Error Mn35: " & Err.Description & " " & Err.Number
    mfGuardaUnPrePago = False
End Function



Public Function mfCierraPrePagos() As Boolean
    On Error GoTo mErrPrePagos3
    If adoPrePag.State = adStateOpen Then adoPrePag.Close
    Set adoPrePag = Nothing
    mfCierraPrePagos = True
    Exit Function
mErrPrePagos3:
    MsgBox "Error Mn36: " & Err.Description & " " & Err.Number
    mfCierraPrePagos = False
End Function



