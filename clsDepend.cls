VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDepend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public adoMome As New ADODB.Recordset
Public vlNroSoc As Long
Public vlDepNum As Long
Public vsDepNmb As String
Public vsDepDoc As String
Public vdDepLim As Single
Public vbDepAut As Boolean






'=====================================================================
Public Sub Inicia()
'=====================================================================
    Set adoMome = New ADODB.Recordset
    Set adoMome.ActiveConnection = adoconn
End Sub



'=====================================================================
Private Sub Class_Terminate()
'=====================================================================
    If adoMome.State = adStateOpen Then
        adoMome.Close
    End If
    Set adoMome = Nothing
    'MsgBox "Termina clase"
End Sub

'=====================================================================
Public Function fBusca2DependientUnSocio() As Boolean
'=====================================================================
    On Error GoTo merr5232
    Dim sMom As String
    If adoMome.State = adStateOpen Then adoMome.Close
    sMom = "SELECT NroSoc, DepNum, DepCI," & _
            "DepNom, DepFechaNac, DepRel," & _
            "DepAuto, DepLimite " & _
            "FROM TBL_Dependientes WHERE NroSoc =" & _
            CStr(vlNroSoc) & " ORDER BY NroSoc;"
    adoMome.Open sMom, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoMome.RecordCount > 0 Then
        fBusca2DependientUnSocio = True
    Else
        fBusca2DependientUnSocio = False
    End If
    Exit Function
    
merr5232:
    MsgBox "ERROR a5232: " & Err.Description & " EN: " & Err.Number
    fBusca2DependientUnSocio = False
End Function



'=====================================================================
Public Function fBuscaDependientUnSocio() As Boolean
'=====================================================================
    On Error GoTo merr5232
    Dim sMom As String
    If adoMome.State = adStateOpen Then adoMome.Close
    sMom = "SELECT  DepNum, DepNom, NroSoc " & _
            "FROM TBL_Dependientes WHERE NroSoc =" & _
            CStr(vlNroSoc) & ";"
    adoMome.Open sMom, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoMome.RecordCount > 0 Then
        fBuscaDependientUnSocio = True
    Else
        fBuscaDependientUnSocio = False
    End If
    Exit Function
    
merr5232:
    MsgBox "ERROR b5232: " & Err.Description & " EN: " & Err.Number
    fBuscaDependientUnSocio = False
End Function


Public Sub fOrdenaAdoPorDepend()
 adoMome("DepNum").Properties("Optimize") = True
'ordena por ese campo
adoMome.Sort = "DepNum"
adoMome.MoveFirst

End Sub
'=====================================================================
Public Function mfBuscaDepend() As Boolean
'=====================================================================
On Error GoTo mError2243
    Dim sCriterio As String
    sCriterio = "DepNum =" & CInt(vlDepNum)
    adoMome.MoveFirst
    adoMome.Find (sCriterio)
    If Not adoMome.EOF Then
        msCamposAVariables
        mfBuscaDepend = True
    Else
        mfBuscaDepend = False
    End If
    'mfBuscaSocio = True
    Exit Function
mError2243:
    MsgBox "ERROR 2243: " & Err.Description & " NE: " & Err.Number
    mfBuscaDepend = False
End Function

'=====================================================================
Public Function mfBusca2Depend() As Boolean
'=====================================================================
'primero busca el socio y luego el depend
adoMome.MoveFirst
adoMome.Find ("NroSoc =" & CLng(vlNroSoc))
If Not adoMome.EOF Then
    adoMome.Find ("DepNum =" & CLng(vlDepNum))
    If adoMome.EOF Then
        mfBusca2Depend = False
        Exit Function
    End If
Else
    mfBusca2Depend = False
    Exit Function
End If
msCamposAVariables
mfBusca2Depend = True
End Function


'=====================================================================
Private Sub msCamposAVariables()
'=====================================================================
vsDepNmb = adoMome("DepNom")
vsDepDoc = adoMome("DepCI")
vdDepLim = adoMome("DepLimite")
vbDepAut = adoMome("depauto")
End Sub

'=====================================================================
Public Function mfAbreTablaDepend() As Boolean
'=====================================================================
    On Error GoTo merr5232
    Dim sMom As String
    If adoMome.State = adStateOpen Then adoMome.Close
    sMom = "SELECT * FROM TBL_Dependientes WHERE NroSoc =" & _
            CStr(vlNroSoc)
    adoMome.Open sMom, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoMome.RecordCount > 0 Then
        mfAbreTablaDepend = True
    Else
        mfAbreTablaDepend = False
    End If
    Exit Function
    
merr5232:
    MsgBox "ERROR a5238: " & Err.Description & " EN: " & Err.Number
    mfAbreTablaDepend = False
End Function


'=====================================================================
Public Function mfAbre2TablaDepend() As Boolean
'=====================================================================
    On Error GoTo merr5233
    Dim sMom As String
    If adoMome.State = adStateOpen Then adoMome.Close
    sMom = "SELECT * FROM TBL_Dependientes ORDER by NroSoc, DepNum;"
    adoMome.Open sMom, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoMome.RecordCount > 0 Then
        mfAbre2TablaDepend = True
    Else
        mfAbre2TablaDepend = False
    End If
    Exit Function
    
merr5233:
    MsgBox "ERROR a5239: " & Err.Description & " EN: " & Err.Number
    mfAbre2TablaDepend = False
End Function
'=====================================================================
Public Sub msTermina()
'=====================================================================
    If adoMome.State = adStateOpen Then adoMome.Close
    Set adoMome = Nothing
End Sub



