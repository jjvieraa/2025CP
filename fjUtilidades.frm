VERSION 5.00
Begin VB.Form fjUtilidades 
   Caption         =   "Utilidades"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Arregla Moneda Pagos"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Revisados:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Conectados:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Sin Socio:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "fjUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoA As New ADODB.Recordset
Dim adoB As New ADODB.Recordset
Dim adoSocio As New ADODB.Recordset
Dim adoNumOrden As New ADODB.Recordset
Dim nCliente As Long
Dim nOrden As Long
Dim nOrdenViejo As Long
Dim nComercio As Long
Dim sCuota As Single
Dim sRecargo As Single
Dim nDato As Byte
Dim Plan As Byte
Dim NroCuota As Byte
Dim dFVto As Date

Dim nM As Long
Dim nM1 As Integer
Dim nM2 As Integer

Private Sub Command1_Click()
adoB.Open "select * FROM tbl_pagos;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
adoB.MoveFirst
Do While Not adoB.EOF
    If Trim(adoB("pag_mon")) = "" Then
        adoB("pag_mon") = "P"
    End If
    adoB.MoveNext
Loop
adoB.MoveFirst
adoB.UpdateBatch adAffectAllChapters
adoB.Close
Set adoB = Nothing
Unload Me
End Sub

Private Sub mome()
Dim sRsp As String
nM = 0
nM1 = 0
nM2 = 0
Me.Show
DoEvents
sRsp = InputBox("Clave para utilidad", "Seguridad")
If Not UCase(sRsp) = "CLAUDI" Then
    Unload Me
    Exit Sub
End If
Screen.MousePointer = vbHourglass

ArreglaMesVto
'cuota          'importacion inicial
'fpruebasuma    'importacion inicial
'importe        'importacion inicial
Screen.MousePointer = vbDefault
End Sub

Private Sub ArreglaMesVto()
Dim dAhora As Date
Dim Final As Date
Dim nMes As Integer
Dim nAño As Integer
Dim ndia As Integer
Dim nCantMeses As Integer

adoB.Open "select * FROM tbl_Ordenes;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
adoB.MoveFirst
Do While Not adoB.EOF
    If adoB!ord_NroOrden > 1519 And _
        adoB!ord_NroOrden < 4225 Then
        nCantMeses = adoB!ord_ctasPagas
        If nCantMeses > 0 Then
            dAhora = adoB!ord_FVto
            nCantMeses = adoB!ord_ctasPagas
            nMes = Month(dAhora)
            nAño = Year(dAhora)
            ndia = Day(dAhora)
            nMes = nMes - nCantMeses
            Do While nMes < 1
                nMes = nMes + 12
                nAño = nAño - 1
            Loop
            Final = CDate(ndia & "/" & nMes & "/" & nAño)
            adoB!ord_FVto = Final
            Debug.Print dAhora & "   " & nCantMeses & "   " & Final
         End If
    End If
    adoB.MoveNext
Loop
adoB.Close
Set adoB = Nothing
End Sub

Private Sub importe()
'importa las ordenes desde las tablas:
adoA.Open "select * FROM orden_2 ORDER BY ord_cli, ord_nro, ord_cta;", adoconn, adOpenStatic, adLockReadOnly, adCmdText
adoB.Open "select * FROM tbl_Ordenes;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

adoA.MoveFirst
nDato = 1
mfTomaNumOrden
fTomaDatosIniciales
Do While Not adoA.EOF
    If Not adoA("ord_cli") = nCliente Or _
        Not adoA("ord_nro") = nOrdenViejo Then
        fGuardaRegistro
        fTomaDatosIniciales
    End If
    fTomaDatos
    Label1.Caption = nM
    Label1.Refresh
    nM = nM + 1
adoA.MoveNext
Loop

msGrabaNumOrden
adoA.Close
Set adoA = Nothing
adoB.Close
Set adoB = Nothing
End Sub


Private Sub fGuardaRegistro()
adoB.AddNew
adoB("ord_nrosoc") = nCliente
adoB("ord_NROCOM") = nComercio
adoB("ord_nroorden") = nOrden
nOrden = nOrden + 1
adoB("ord_Cuota") = sCuota
adoB("ord_plan") = Plan
adoB("ord_ctasPagas") = NroCuota
adoB("ord_Femis") = CDate("01/01/1900")    'cambiar
adoB("ord_fvto") = dFVto
adoB("ord_cerro") = CDate("01/01/1900")    'cambiar
adoB("ord_mon") = "P"
adoB("ord_recarg") = sRecargo
adoB("ord_fHora") = nOrdenViejo
adoB.Update
Label2.Caption = nM1
Label2.Refresh
nM1 = nM1 + 1
End Sub
Private Sub fTomaDatosIniciales()
Plan = 0
nDato = 1
sRecargo = 0
nCliente = CLng(adoA("ord_cli"))
nComercio = CLng(adoA("ord_com"))
sCuota = adoA("ord_val")
nOrdenViejo = CLng(adoA("ord_nro"))
dFVto = CDate("10/" & adoA("ord_mes") & "/" & adoA("ord_ano"))
End Sub

Private Sub fTomaDatos()
Debug.Print adoA("ord_cli") & "  " & adoA("ord_nro")
If nDato = 1 Then
    If adoA("ord_cta") = 1 Then
        NroCuota = 0
        Plan = 1
    Else
        NroCuota = adoA("ord_cta") - 1
        Plan = adoA("ord_cta")
    End If
Else
    Plan = Plan + 1
End If
sRecargo = sRecargo + adoA("ord_sal") - adoA("ord_val")
nDato = nDato + 1
End Sub



Public Sub mfTomaNumOrden()
    
    adoNumOrden.Open "SELECT * FROM TBL_Parametros", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    adoNumOrden.MoveFirst
    nOrden = adoNumOrden("NroOrden") + 1
     adoNumOrden.Close
    Set adoNumOrden = Nothing
End Sub
Public Sub msGrabaNumOrden()
    
    adoNumOrden.Open "SELECT * FROM TBL_Parametros", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    adoNumOrden.MoveFirst
    adoNumOrden("NroOrden") = nOrden
    adoNumOrden.Update
    adoNumOrden.Close
    Set adoNumOrden = Nothing
End Sub


Private Sub cuota()
'importa las ordenes desde las tablas:
adoA.Open "select * FROM sit_cli;", adoconn, adOpenStatic, adLockReadOnly, adCmdText
adoB.Open "select * FROM tbl_Ordenes;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
adoSocio.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

adoA.MoveFirst
nDato = 1
Do While Not adoA.EOF
    If adoA("sit_val") > 0 Then
        fGuarda2Registro
    End If
    If adoA("sit_otros") > 0 Then
        fGuarda3Registro
    End If
    Label1.Caption = nM
    Label1.Refresh
    nM = nM + 1
    adoA.MoveNext
Loop

adoA.Close
Set adoA = Nothing
adoB.Close
Set adoB = Nothing
adoSocio.Close
Set adoSocio = Nothing
End Sub

Private Sub fGuarda2Registro()
If NoExisteSocio Then Exit Sub
adoB.AddNew
adoB("ord_nrosoc") = CLng(adoA("sit_cli"))
adoB("ord_NROCOM") = "0"
adoB("ord_nroorden") = 1
adoB("ord_Cuota") = adoA("sit_val")
adoB("ord_plan") = 1
adoB("ord_ctasPagas") = 0
adoB("ord_Femis") = CDate("01/01/1900")    'cambiar
adoB("ord_fvto") = CDate("10/" & adoA("sit_mes") & "/" & adoA("sit_ano"))
adoB("ord_cerro") = CDate("01/01/1900")    'cambiar
adoB("ord_mon") = "P"
adoB("ord_recarg") = 0
adoB("ord_fHora") = "sv"
adoB.Update
End Sub
Private Sub fGuarda3Registro()
If NoExisteSocio Then Exit Sub
adoB.AddNew
adoB("ord_nrosoc") = CLng(adoA("sit_cli"))
adoB("ord_NROCOM") = "0"
adoB("ord_nroorden") = 2
adoB("ord_Cuota") = adoA("sit_otros")
adoB("ord_plan") = 1
adoB("ord_ctasPagas") = 0
adoB("ord_Femis") = CDate("01/01/1900")    'cambiar
adoB("ord_fvto") = CDate("10/" & adoA("sit_mes") & "/" & adoA("sit_ano"))
adoB("ord_cerro") = CDate("01/01/1900")    'cambiar
adoB("ord_mon") = "P"
adoB("ord_recarg") = 0
adoB("ord_fHora") = "sv"
adoB.Update
End Sub

Private Function NoExisteSocio() As Boolean
    adoSocio.MoveFirst
    adoSocio.Find ("NroSoc =" & adoA("sit_cli"))
    If Not adoSocio.EOF Then
        Label6.Caption = nM2
        Label6.Refresh
        nM2 = nM2 + 1
        NoExisteSocio = False
    Else
        Label2.Caption = nM1
        Label2.Refresh
        nM1 = nM1 + 1
        NoExisteSocio = True
    End If
End Function
Private Sub fpruebasuma()
Dim sMome As String
sMome = "SELECT SUM(sit_val) as a1, sum(sit_otros) as a2 fROM sit_cli;"
adoA.Open sMome, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
MsgBox adoA!a1 & "  " & adoA!a2
adoA.Close
Set adoA = Nothing
End Sub

