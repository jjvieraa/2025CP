VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim adoSocios As New ADODB.Recordset
Public vlNroSoc As Long       'NroSoc
Public vsNroCob As String       'NroCob  11
Public vdFIng As Date         'Fech_Ing
Public vsApellido As String   '= 25
Public vsNombre As String     '= 25
Public vsDireccion As String  ' = 50
Public vsLocalidad As String  ' = 30
Public vsTel As String        ' = 25
Public vdFNac As Date         ' Fech_Nac
Public vlCodCatSoc As Long
Public vlCodSitLab As Long
Public vlCodUnidPer As Long
Public vlCodPresServ As Long
Public vlCodGrado As Long
Public vlCodEstCiv As Long
Public vsCi As String         ' = 11
Public vsOcupacion As String  ' =15
Public vbAyuda As Boolean
Public vnCobrador As Byte
Public vnsIngresos As Single
Public vbCredAutorizado As Boolean 'Cred_Auto
Public vnsLimite As Single
Public vnlNroGarant As Long        'Garantia
Public vsCOP As String            ' = 2
Public vsFuncionario As String    ' = 2
Public vsFuncDia  As String       ' FDia 10
Public vsFuncHora As String         ' FHora 8
Public vsSolicit As String   'Numero solicitud = 12



'=====================================================================
Public Sub msInicia()
'=====================================================================
    Set adoSocios = New ADODB.Recordset
    Set adoSocios.ActiveConnection = adoConn
End Sub


'=====================================================================
Public Sub msTermina()
'=====================================================================
    If adoSocios.State = adStateOpen Then adoSocios.Close
    Set adoSocios = Nothing
End Sub


'=====================================================================
Public Function mfBuscaSocio() As Boolean
'=====================================================================
On Error GoTo mError2235
    Dim sCriterio As String
    sCriterio = "NroSoc =" & CInt(vlNroSoc)
    adoSocios.MoveFirst
    adoSocios.Find (sCriterio)
    If Not adoSocios.EOF Then
        msCamposAVariables
        mfBuscaSocio = True
    Else
        mfBuscaSocio = False
    End If
    'mfBuscaSocio = True
    Exit Function
mError2235:
    MsgBox "ERROR 2235: " & Err.Description & " NE: " & Err.Number
    mfBuscaSocio = False
End Function

'=====================================================================
Public Function mfAbreTablaSociosOrdenSocio() As Boolean
'=====================================================================
On Error GoTo mErr2234
    If adoSocios.State = adStateOpen Then adoSocios.Close
    adoSocios.Open "SELECT * FROM TBL_Socios ORDER by NroSoc;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    mfAbreTablaSociosOrdenSocio = True
    Exit Function
mErr2234:
    MsgBox "ERROR 2234: " & Err.Description & " NE: " & Err.Number
    mfAbreTablaSociosOrdenSocio = False
End Function

'=====================================================================
Private Sub msVariablesACampos()
'=====================================================================
adoSocios!NroSoc = vlNroSoc
adoSocios!NroCob = vsNroCob
adoSocios!Fech_ing = vdFIng
adoSocios!Apellido = vsApellido
adoSocios!nombre = vsNombre
adoSocios!direccion = vsDireccion
adoSocios!localidad = vsLocalidad
adoSocios!Tel = vsTel      ' POR ALGUNOS QUE ESTAN VACIOS
adoSocios!Fech_nac = vdFNac
adoSocios!CodCatSoc = vlCodCatSoc
adoSocios!CodSitLab = vlCodSitLab
adoSocios!codunidper = vlCodUnidPer
adoSocios!codpresserv = vlCodPresServ
adoSocios!CodGrado = vlCodGrado
adoSocios!CodEstCiv = vlCodEstCiv
adoSocios!ci = vsCi
adoSocios!ocupacion = vsOcupacion
adoSocios!ayuda = vbAyuda
adoSocios!cobrador = vnCobrador
adoSocios!ingresos = vnsIngresos
adoSocios!Cred_Auto = vbCredAutorizado
adoSocios!Limite = vnsLimite
adoSocios!Garantia = vnlNroGarant
adoSocios!COP = vsCOP
adoSocios!Funcionario = vsFuncionario
adoSocios!FDia = vsFuncDia
adoSocios!FHora = vsFuncHora
adoSocios!Solicitud = vsSolicit
End Sub

'=====================================================================
Private Sub msCamposAVariables()
'=====================================================================

vlNroSoc = adoSocios!NroSoc
vsNroCob = adoSocios!NroCob
vdFIng = adoSocios!Fech_ing
vsApellido = adoSocios!Apellido
vsNombre = adoSocios!nombre
vsDireccion = "" & adoSocios!direccion
vsLocalidad = "" & adoSocios!localidad
vsTel = "" & adoSocios!Tel
vdFNac = adoSocios!Fech_nac
vlCodCatSoc = adoSocios!CodCatSoc
vlCodSitLab = adoSocios!CodSitLab
vlCodUnidPer = adoSocios!codunidper
vlCodPresServ = adoSocios!codpresserv
vlCodGrado = adoSocios!CodGrado
vlCodEstCiv = adoSocios!CodEstCiv
vsCi = adoSocios!ci
vsOcupacion = "" & adoSocios!ocupacion
vbAyuda = adoSocios!ayuda
vnCobrador = adoSocios!cobrador
vnsIngresos = adoSocios!ingresos
vbCredAutorizado = adoSocios!Cred_Auto
vnsLimite = adoSocios!Limite
vnlNroGarant = adoSocios!Garantia
vsCOP = "" & adoSocios!COP
vsFuncionario = "" & adoSocios!Funcionario
vsFuncDia = "" & adoSocios!FDia
vsFuncHora = "" & adoSocios!FHora
vsSolicit = "" & adoSocios!Solicitud
End Sub
