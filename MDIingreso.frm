VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIingreso 
   BackColor       =   &H8000000C&
   Caption         =   "Círculo Policial"
   ClientHeight    =   6870
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6660
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIingreso.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6615
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1766
            MinWidth        =   1766
            TextSave        =   "14/11/2010"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "hhh"
            TextSave        =   "hhh"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1412
            MinWidth        =   1412
            TextSave        =   "NÚM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu arch 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu mnuclientes 
         Caption         =   "&Socios"
         Begin VB.Menu ingsoc 
            Caption         =   "&Ingresar"
         End
         Begin VB.Menu modsoc 
            Caption         =   "&Modificar"
         End
         Begin VB.Menu buscsoc 
            Caption         =   "&Buscar"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuDepe 
         Caption         =   "&Dependientes"
      End
      Begin VB.Menu mnuComercios 
         Caption         =   "&Comercios"
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVarios 
         Caption         =   "&Varios"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu nmContra 
         Caption         =   "C&ontraseña"
      End
      Begin VB.Menu nmcontroles 
         Caption         =   "Co&ntroles"
         Begin VB.Menu nmOYC 
            Caption         =   "Orden y Comercios"
         End
         Begin VB.Menu mnuOD 
            Caption         =   "Ordenes Dobles"
         End
      End
      Begin VB.Menu p 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEspecial 
         Caption         =   "Especial3"
      End
      Begin VB.Menu G1 
         Caption         =   "&G1"
      End
      Begin VB.Menu G2 
         Caption         =   "G2"
      End
      Begin VB.Menu mInfoAdm 
         Caption         =   "InfoAdm"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mCaja 
      Caption         =   "&Caja"
      Begin VB.Menu mIngresos 
         Caption         =   "&Ingresos"
      End
      Begin VB.Menu mEgresos 
         Caption         =   "&Ent y Sal"
      End
      Begin VB.Menu Admin 
         Caption         =   "&Admin"
      End
      Begin VB.Menu mAP 
         Caption         =   "&Anula Pago"
      End
   End
   Begin VB.Menu mOperciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mOrden 
         Caption         =   "&Orden"
         Begin VB.Menu mIngresoOrden 
            Caption         =   "&Ingreso Orden"
         End
         Begin VB.Menu mAnular 
            Caption         =   "&Anular"
         End
         Begin VB.Menu mRecorrer 
            Caption         =   "&Recorrer"
         End
         Begin VB.Menu mImprimir 
            Caption         =   "I&mprimir"
         End
      End
      Begin VB.Menu mRecibos 
         Caption         =   "&Recibos"
      End
      Begin VB.Menu mnComercios 
         Caption         =   "C&omercios"
      End
      Begin VB.Menu mnCierre 
         Caption         =   "&Cierre"
      End
      Begin VB.Menu mInfo 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mAcerca 
         Caption         =   "&Acerca"
      End
   End
   Begin VB.Menu mnSalir 
      Caption         =   "&Salir"
   End
   Begin VB.Menu mFuncionario 
      Caption         =   "&Func"
   End
End
Attribute VB_Name = "MDIingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents rstestciv As Recordset
Attribute rstestciv.VB_VarHelpID = -1
Dim adoM As New ADODB.Recordset






Private Sub mAP_Click()
    If vpnNivelFuncionario < kNivel3 Then Exit Sub
    fjCobroAnula.Show
End Sub

'===========================================
'CONSTANTES DE COMPILACION CONDICIONAL:
'kCASA en modulo fjORDENES
'
'
'===========================================

Private Sub MDIForm_Load()
Dim mComando As ADODB.Command

On Error GoTo mMal


'EVITA QUE SE EJECUTE 2 VECES
If App.PrevInstance Then
    MsgBox "Este programa YA se está ejecutando"
    End
End If
vpbVieneDeMuestraTabla = False
Set mComando = New ADODB.Command
On Error Resume Next
'conectar
Set adoConn = New ADODB.Connection
adoConn.CursorLocation = adUseClient
adoConn.Open "dsn=jimmy;Jet OLEDB:Database Password=cp7;"""
'MsgBox "Conexión satisfactoria" & Chr(13) & " Ingresando al Sistema... ", vbInformation, "¡Bienvenido!"

' Pide contraseña
frmjLogin.Show vbModal       'JV
DoEvents
Me.Show
MouseOff

fjAviso3.Label1.Caption = "Espere..."
fjAviso3.Show
DoEvents
'cargar datos
StatusBar1.Panels(2).Text = "Cargando datos..."

mComando.ActiveConnection = adoConn
mComando.CommandType = adCmdText
Set rstestciv = New Recordset
rstestciv.Open "select * from estcivil where idEstCiv = 0", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If rstestciv.RecordCount = 0 Then   'SOLO PARA EL PRINCIPIO QUE NO HAY VALORES
        rstestciv.Close
        mComando.CommandText = "insert into estcivil values (0, '(Ninguno)')"
        mComando.Execute
        
        mComando.CommandText = "insert into grado values (0, '(Ninguno)')"
        mComando.Execute
        
        mComando.CommandText = "insert into catsocio values (0, '(Ninguna)')"
        mComando.Execute

        mComando.CommandText = "insert into slaboral values (0, '(Ninguna)')"
        mComando.Execute

        mComando.CommandText = "insert into unidadpert values (0, '(Ninguna)')"
        mComando.Execute
        
        mComando.CommandText = "insert into unidadserv values (0, '(Ninguna)')"
        mComando.Execute
            
        mComando.CommandText = "insert into rubro values (0, '(Ninguno)')"
        mComando.Execute
    
End If
Set rstestciv = Nothing
Set mComando = Nothing

' Parametros
StatusBar1.Panels(2).Text = "Leyendo parámetros..."
StatusBar1.Refresh
sLeeParamPublicos

' Abriendo la tabla Ordenes, globalmente
DoEvents
StatusBar1.Panels(2).Text = "ESPERE: Abriendo Ordenes..."
StatusBar1.Refresh
'Set mGlob.cOrd = New clsOrdenes
mGlob.cOrd.msAbreTablaOrdenesGlobal
'mGlob.cOrd.msOrdenaTablaOrdenesGlobalPorNroOrden

' Abriendo la tabla pagos, globalmente
DoEvents
StatusBar1.Panels(2).Text = "ESPERE: Abriendo Pagos..."
StatusBar1.Refresh
If Not mGlob.cPgs.mfAbrePagos Then
    MsgBox "Error 3748: al abrir la tabla Pagos. No: " & Err.Number & _
         vbNewLine & Err.Description, vbCritical, "Problemas"
    Unload Me
    End
End If

DoEvents
StatusBar1.Panels(2).Text = ""
StatusBar1.Refresh
fjAviso3.Hide
MouseOn
Exit Sub
'otro

mMal:
   MsgBox "Error 3100: al inicializar archivos. No: " & Err.Number & _
         vbNewLine & Err.Description, vbCritical, "Problemas"
         End

End Sub





'======================================
'entrada a los programas
'======================================


Private Sub mAcerca_Click()
fjAcerca.Show
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'jv cierra la historia de un funcionario
HistoriaSale
'jv limpia memoria
adoConn.Close
Set adoConn = Nothing

Set cOrd = Nothing
Set cPgs = Nothing
If adoM.State = adStateOpen Then adoM.Close
Set adoM = Nothing
If adoOrdns.State = adStateOpen Then adoOrdns.Close
Set adoOrdns = Nothing
End Sub

'Private Sub CierraTodo()
'adoConn.Close
'Set adoConn = Nothing
'End'
'
'End Sub

Private Sub mnuAyuda_Click()
MsgBox "Ayuda no Disponible", vbOKOnly, "Ayuda"
End Sub



Private Sub mOrdenes_Click()
If vpnNivelFuncionario < kNivel2 Then Exit Sub
fjIngComercio.Show
End Sub

'================================
Private Sub ingsoc_Click()
'ingreso socios
If vpnNivelFuncionario < kNivel1 Then Exit Sub
vpFormMovim = kFormIngresa
'fjIngresos.Show
fjIngresos2.Show
End Sub




Private Sub mFuncionario_Click()
frmjLogin.Show vbModal       'JV
End Sub

Private Sub mInfo_Click()
If vpnNivelFuncionario < kNivel2 Then Exit Sub
fjInformes.Show
End Sub

Private Sub mIngresoOrden_Click()
If vpnNivelFuncionario < kNivel3 Then Exit Sub
'
'DoEvents
fjOrdenes.Show
End Sub


Private Sub mIngresos_Click()
fjCobros.Show
End Sub

Private Sub mInfoAdm_click()
If vpnNivelFuncionario < kNivel6 Then Exit Sub
frmMisInformes.Show
End Sub

Private Sub mMome_Click()
If vpnNivelFuncionario < kNivel6 Then Exit Sub
fjUtilidades.Show
End Sub

Private Sub mnCierre_Click()
'si el usuario tiene autorizacion
If vpnNivelFuncionario < kNivel5 Then Exit Sub
    fjCierraMes.Show
End Sub

Private Sub mnnuUtilidad_Click()
If vpnNivelFuncionario < kNivel6 Then Exit Sub
fjUtilidades.Show
End Sub

Private Sub mnComercios_Click()
fjComercios.Show
End Sub

Private Sub mnuEspec2_Click()
If vpnNivelFuncionario < kNivel6 Then Exit Sub
fjMome2.Show

End Sub

Private Sub mnuEspecial_Click()
If vpnNivelFuncionario < kNivel6 Then Exit Sub
fjMome2.Show
End Sub

Private Sub mnuOD_Click()
'busca ordenes dobles
Dim nM As Long
Dim nM1 As Long
Screen.MousePointer = vbHourglass
 
nM1 = 0
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "select * from tbl_ordenes ORDER BY ord_NroOrden", adoConn, adOpenDynamic, adLockOptimistic
adoM.MoveFirst
nM = 0
Do While Not adoM.EOF
    If adoM!ord_NroOrden > 2 And adoM!ord_NroOrden = nM Then
        MsgBox "Orden Nro " & nM & " está doble"
        nM1 = nM1 + 1
    End If
    nM = adoM!ord_NroOrden
    adoM.MoveNext
Loop
If nM1 = 0 Then
    MsgBox "Sin órdenes dobles"
Else
    MsgBox "Hay " & nM1 & " ordenes dobles"
End If
If adoM.State = adStateOpen Then adoM.Close
Set adoM = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub modsoc_Click()
'Modifica SOCIo
If vpnNivelFuncionario < kNivel2 Then Exit Sub
fjMantenSocio.Show
End Sub
Private Sub mnuDepe_Click()
'DEPENDIENTES
If vpnNivelFuncionario < kNivel2 Then Exit Sub
fjMantenDepend.Show
End Sub
Private Sub mnuComercios_Click()
'COMERCIOS
If vpnNivelFuncionario < kNivel2 Then Exit Sub
fjIngComercio.Show
End Sub

Private Sub mnuVarios_Click()
'MANEJO DE PARAMETROS
If vpnNivelFuncionario < kNivel5 Then Exit Sub
fjDatos.Show
End Sub
Private Sub mnSalir_Click()
        Unload Me
        End
End Sub

Private Sub mRecibos_Click()
If vpnNivelFuncionario < kNivel2 Then Exit Sub
fjRecibo.Show
End Sub

Private Sub mRecorrer_Click()
If vpnNivelFuncionario < kNivel2 Then Exit Sub
vpFormMovim = kFormRecorre
fjOrden2.Show
End Sub

Private Sub mAnular_Click()
If vpnNivelFuncionario < kNivel3 Then Exit Sub
fjPideDato.Caption = "Nro de Orden:"
fjPideDato.Show vbModal
If vpbCancel = False Then
    vpFormMovim = kFormAnula
    fjOrden2.Show
End If
End Sub





'==========================================================
Private Sub sLeeParamPublicos()
'==========================================================
    'mAviso ("Leyendo Parametros...")
    If Not mGlob.fTomaMesOperacYParametros Then
        End
        Exit Sub
    End If

    'verifico la el año de Presupuesto este en el intervalo correcto
     Dim dMome As Integer
     dMome = Year(CDate(CStr(vpnPrspHst) & "/" & _
        Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)))
     If dMome > kMayorAñoTrabajo Or _
        dMome < kMenorAñoTrabajo Then
        mfAviso ("Año de presupuesto fuera de intervalo.")
        Unload Me
        End
     End If
     
       
End Sub


Private Sub nmContra_Click()
'CAMBIA LA CONTRASEÑA DEL USUARIO ACTUAL
Dim sRsp As String
Dim sRsp2 As String
sRsp = InputBox("Indique su contraseña actual", "Cambio de contraseña")
If Not UCase(sRsp) = UCase(vptFuncPass) Then
    Exit Sub
End If
sRsp = InputBox("Indique su contraseña nueva: ", "Cambio de contraseña")
sRsp2 = InputBox("Repitala: ", "Cambio de contraseña")
If Not UCase(sRsp) = UCase(sRsp2) Then
    MsgBox "No hubo cambio", vbCritical, "Cambio de contraseña"
Else
    If CambioDeContraseña(UCase(sRsp)) Then
        MsgBox "Cambio Realizado", vbInformation, "Cambio de contraseña"
    Else
        MsgBox "No hubo cambio", vbCritical, "Cambio de contraseña"
    End If
End If
End Sub

Private Sub nmOYC_Click()
'busca ordenes sin comercio

Dim nM As Long
Dim nM1 As Long
             Screen.MousePointer = vbHourglass
 
  If adoM.State = adStateOpen Then adoM.Close
     adoM.Open "select * from tbl_ordenes", adoConn, adOpenDynamic, adLockOptimistic
nM = adoM.RecordCount
 If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "select * from tbl_ordenes INNER JOIN tbl_comercios ON tbl_comercios.codigo = tbl_ordenes.ord_nrocom;", adoConn, adOpenDynamic, adLockOptimistic
nM1 = adoM.RecordCount
 If adoM.State = adStateOpen Then adoM.Close
 Set adoM = Nothing
If Not nM = nM1 Then MsgBox "Hay " & nM - nM1 & " ordenes sin relación con comercios"
            Screen.MousePointer = vbDefault
End Sub

' Caja > Egresos
Private Sub mEgresos_Click()
    vpMuestraTabla = 1         'Ent y Sal funcionarios
    fjSalidasYEntradas.Show
End Sub


'Programa de entradas y salidas de administrador del CP
Private Sub Admin_Click()
If Not vpnFuncionario = 5 Then Exit Sub 'solo para el administrador
    vpMuestraTabla = 2     'Ent y Sal Administrador
    fjSalidasYEntradas.Show
End Sub


'puede modificar las entradas y salidas
Private Sub G1_Click()
If Not (vpnFuncionario = 5 Or vpnFuncionario = 30) Then Exit Sub 'solo para el administrador
vpMuestraTabla = 1
fjModifGastos.Show
End Sub

'puede modificar las entradas y salidas del administrador
Private Sub G2_Click()
If Not (vpnFuncionario = 5 Or vpnFuncionario = 30) Then Exit Sub 'solo para el administrador
vpMuestraTabla = 2
fjModifGastos.Show
End Sub


