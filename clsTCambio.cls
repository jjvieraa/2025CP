VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public adoTC As ADODB.Recordset

Public vdFecha As Date
Public vnsDolar As Single
Public vnsReal As Single
Public vnsPArg As Single
Public vnsUR As Single


Public vnbQueMoneda As Byte
Const vkDolar = 1
Const vkPArg = 2
Const vkReal = 3
Const vkUR = 4

Const vksPesos = "P"
Const vksDolar = "D"
Const vksPArg = "A"
Const vksReal = "R"
Const vksUR = "U"
  
  

'=====================================================================
Public Sub msInicia()
'=====================================================================
    On Error GoTo merr52332
    Set adoTC = New ADODB.Recordset
    Set adoTC.ActiveConnection = adoconn
    If adoTC.State = adStateOpen Then adoTC.Close
    adoTC.Open "SELECT * FROM tbl_TCambio ORDER BY TC_Fecha;", adoconn, adOpenDynamic, adLockOptimistic
    Exit Sub
    
merr52332:
    MsgBox "ERROR 52332: " & Err.Description & " EN: " & Err.Number
End Sub


'=====================================================================
Public Sub msTermina()
'=====================================================================
    If adoTC.State = adStateOpen Then
        adoTC.Close
    End If
    Set adoTC = Nothing
End Sub


'=====================================================================
Public Sub msInicializaVariables()
'=====================================================================
    'vdFecha = "00/00/00"
    vnsDolar = 0#
    vnsReal = 0#
    vnsPArg = 0#
    vnsUR = 0#
End Sub

'=====================================================================
Private Sub msVariablesACampos()
'=====================================================================

    adoTC!TC_Fecha = vdFecha
    adoTC!TC_Dolar = vnsDolar
    adoTC!TC_Real = vnsReal
    adoTC!TC_PArg = vnsPArg
    adoTC!TC_UR = vnsUR
End Sub
'=====================================================================
Private Sub msCamposAVariables()
'=====================================================================
    vdFecha = adoTC!TC_Fecha
    vnsDolar = mfEsNulo(adoTC!TC_Dolar)
    vnsReal = mfEsNulo(adoTC!TC_Real)
    vnsPArg = mfEsNulo(adoTC!TC_PArg)
    vnsUR = mfEsNulo(adoTC!TC_UR)
End Sub


'=====================================================================
Public Function mfBuscaTCambio() As Single
'=====================================================================
'vnbQueMoneda indica la moneda que se quiere pedir
On Error GoTo mError22a35

    Dim sCriterio As String
    Dim sMome As String
    
    If vnbQueMoneda = 0 Then
        mfBuscaTCambio = 0#
        Exit Function
    End If
   ' sMome = mfInvierteMes(CStr(vdFecha))
    sMome = CStr(vdFecha)
    sCriterio = "TC_Fecha =#" & sMome & "#"
OtraVez:
    adoTC.MoveFirst
    adoTC.Find (sCriterio)
    If Not adoTC.EOF Then
        vplNumeroRegistro = adoTC.Bookmark
        msCamposAVariables
        'Busca que tenga valor esa monead
        Select Case vnbQueMoneda
            Case vkDolar
                If vnsDolar = 0# Then msPideTCambios
             Case vkPArg
                If vnsPArg = 0# Then msPideTCambios
            Case vkReal
                If vnsReal = 0# Then msPideTCambios
            Case vkUR
                If vnsUR = 0# Then msPideTCambios
        End Select
    Else            'pide valores
        'AGREGA REGISTRO
        adoTC.AddNew
        vplNumeroRegistro = adoTC.Bookmark
        adoTC!TC_Fecha = vdFecha
        adoTC!TC_Dolar = 0#
        adoTC!TC_Real = 0#
        adoTC!TC_PArg = 0#
        adoTC!TC_UR = 0#
        adoTC.Update
        GoTo OtraVez
    End If
    Select Case vnbQueMoneda
        Case vkDolar
            mfBuscaTCambio = vnsDolar
        Case vkPArg
            mfBuscaTCambio = vnsPArg
        Case vkReal
            mfBuscaTCambio = vnsReal
        Case vkUR
            mfBuscaTCambio = vnsUR
        Case Else
            mfBuscaTCambio = 0#
    End Select
    Exit Function
mError22a35:
    mMsgErr "ERROR 22a35: " & Err.Description & " NE: " & Err.Number
End Function

'=====================================================================
Public Sub msPideTCambios()
'=====================================================================
'vnsDolar = InputBox("Dolar:", "Tasa de Cambio del")
fjPideTCbio.Caption = "Tasa de Cambio del " & vdFecha
fjPideTCbio.Text1(0).Text = vnsDolar
fjPideTCbio.Text1(1).Text = vnsPArg
fjPideTCbio.Text1(2).Text = vnsReal
fjPideTCbio.Text1(3).Text = vnsUR
fjPideTCbio.Show vbModal

'regresó del fjPideTCbio
fjPideTCbio.Hide
fjPideTCbio.Refresh
vnsDolar = vpTCD
vnsReal = vpTCR
vnsPArg = vpTCA
vnsUR = vpTCU
Me.msGuardaTCambios
End Sub

'=====================================================================
Public Sub msGuardaTCambios()
'=====================================================================
adoTC.MoveFirst
adoTC.Bookmark = vplNumeroRegistro
'adoTC.Move vplNumeroRegistro
'adoTC!TC_Fecha = vdFecha
adoTC!TC_Dolar = vnsDolar
adoTC!TC_Real = vnsReal
adoTC!TC_PArg = vnsPArg
adoTC!TC_UR = vnsUR
adoTC.Update
End Sub

Public Sub msDeterminaMoneda(sPrm As String)
Select Case sPrm
    Case "A"
        vnbQueMoneda = 2
    Case "R"
        vnbQueMoneda = 3
    Case "D"
        vnbQueMoneda = 1
    Case "U"
        vnbQueMoneda = 4
    Case Else
        vnbQueMoneda = 0
End Select
End Sub


Public Function mfDevuelveCambio(sMon As String, dFech As Date) As Single
        Dim sMome As Single
        
        
        Me.msInicia
        Me.msInicializaVariables
        Me.vdFecha = dFech     ' fecha de hoy
        Me.msDeterminaMoneda (sMon)
        sMome = Me.mfBuscaTCambio
        Me.msTermina
        mfDevuelveCambio = sMome
End Function

