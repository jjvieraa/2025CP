VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComercios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit
Dim rstM As ADODB.Recordset
Dim nOrden() As Integer
Dim nTope As Integer
Public nbCantRec As Byte

Public Function CargaGruposEnCombo(cPrm As ComboBox) As Boolean
    Dim i As Integer
    
    Set rstM = New Recordset
    'Abre los grupos de Comercios ordenados por nombre(Farmacias, provisiones, etc)
    rstM.Open "SELECT * FROM tbl_Grupos ORDER BY deta;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstM.RecordCount <> 0 Then
        'determina la cantidad de registros
        nTope = rstM.RecordCount
        ReDim nOrden(nTope) As Integer
        'recorre el recordset
        For i = 0 To nTope - 1
            'guarda los grupos en el combo
            cPrm.AddItem ("" & rstM!Deta)
            'guarda los numero de grupo en la matriz
            nOrden(i) = rstM!idrubro
            
            rstM.MoveNext
        Next i
        CargaGruposEnCombo = True
    Else
        CargaGruposEnCombo = False
    End If
    rstM.Close
    Set rstM = Nothing
End Function


Public Function DevuelveNoRubro(nPrm As Integer) As Integer
DevuelveNoRubro = nOrden(nPrm)
End Function


Public Function DevuelveNoDelCombo(nPrm As Integer) As Integer
Dim ni As Integer
For ni = 0 To nTope - 1
    If nOrden(ni) = nPrm Then
        DevuelveNoDelCombo = ni
        Exit Function
    End If
Next
DevuelveNoDelCombo = 0
End Function


Public Function BuscaComercio(lPrm As Long) As String
    'Busca el comercio por el CODIGO (Nro de comercio) en la tbl_Comercios 
    Set rstM = New Recordset
    rstM.Open "SELECT * FROM tbl_comercios ORDER BY codigo;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    'no tiene registros
    If rstM.RecordCount < 1 Then
        BuscaComercio = "Sin registros..."
        Exit Function
    End If
    rstM.Find ("codigo =" & lPrm)
    'no encontro el COMERCIO con ese numero
    If rstM.EOF Then
        BuscaComercio = "No encontrado..."
        Exit Function
    End If
    ' guarda en la variable nbCantRec el número que tiene en el campo CantRecib "Cantidad de recibos a imprimir"
    nbCantRec = rstM("CantRecib")
    'Devuelve el Nombre del Comercio
    BuscaComercio = rstM("NombCom")
End Function


'=====================================================================
Public Function mfAbreTablaComercios() As Boolean
'=====================================================================
On Error GoTo mErr2234
    Set rstM = New Recordset
    If rstM.State = adStateOpen Then rstM.Close
    rstM.Open "SELECT * FROM tbl_comercios ORDER BY codigo;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    mfAbreTablaComercios = True
    Exit Function
mErr2234:
    MsgBox "ERROR b2234: " & Err.Description & " NE: " & Err.Number
    mfAbreTablaComercios = True
End Function


Public Function BuscaComercio2(lPrm As Long) As String
'busca el compercio por el CODIGO (nro) y devuelve el NOMBRE
    If lPrm = 0 Then
        BuscaComercio2 = ""
        Exit Function
    End If
    rstM.MoveFirst
    rstM.Find ("codigo =" & lPrm)
    If rstM.EOF Then
        BuscaComercio2 = "No encontrado..."
        Exit Function
    End If
    BuscaComercio2 = rstM("NombCom")
End Function
