VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fjCierraMes 
   BackColor       =   &H8000000D&
   Caption         =   "Cierra Mes"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fjCierraMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdArchivoJefat 
      BackColor       =   &H00FF8080&
      Caption         =   "Archivo Jefat"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "crea c:\cp\R521.txt"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdJefListaAlf 
      BackColor       =   &H00FF8080&
      Caption         =   "L. Jefat. Alfab."
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdArchivCentro 
      BackColor       =   &H00FF8080&
      Caption         =   "Archivo Centro"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "crea c:\cp\R521.txt"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdListaCentro 
      BackColor       =   &H00FF8080&
      Caption         =   "Listado Centro"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcedidos 
      BackColor       =   &H00FF8080&
      Caption         =   "Excedidos"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdJefLista 
      BackColor       =   &H00FF8080&
      Caption         =   "Listado  Jefat"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCentroPol 
      BackColor       =   &H00FF8080&
      Caption         =   "Disq.Centro"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdJefat 
      BackColor       =   &H00FF8080&
      Caption         =   "Disq.Jefatura"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdResumen 
      BackColor       =   &H00FF8080&
      Caption         =   "Resumen"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FF8080&
      Caption         =   "Salir"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCierraMes 
      BackColor       =   &H00FF8080&
      Caption         =   "Cierra Mes"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdParametros 
      BackColor       =   &H00FF8080&
      Caption         =   "Parámetros"
      Height          =   255
      Left            =   3600
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblOtro 
      BackColor       =   &H8000000D&
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lblDeta 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label lblMes 
      BackColor       =   &H8000000D&
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "fjCierraMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CIERRA MES: genera los registros en tbl_prepago y genera registros por cuota y por ayuda en tbl_ordenes
' RESUMEN: lista el tbl_prepago

' CENTRO: A) DISQUETE:             prepago --> DTO
' CENTRO: B) LISTADO :                prepago -->
' CENTRO: C) ARCHIVO:                prepago --> r521.txt

' JEFAT: A) DISQUETE:                  prepago --> jefatura --> jef3  --> jefat.xls
' JEFAT: B) LISTADO:                     jefatura -->
' JEFAT: C) LIST. ALFAB:              jefatura  -->
' JEFAT: D) ARCHIVO:                    prepago  --> raamm.txt

'=======================================================================
Option Explicit
Dim nEx As Object
Dim o_Hoja  As Object

Dim adoM As New ADODB.Recordset
Dim adoP As New ADODB.Recordset
Dim adoClie As New ADODB.Recordset
Dim adoCom As New ADODB.Recordset
Dim adoOrd As New ADODB.Recordset

Dim dFechaVtoActual As Date

Dim clsOrd As New clsOrdenes
Dim clsPag As New clsPagos
Dim clsPreP As New clsPrePag
Dim clsRecib As New clsRecibos

Dim mNroOrden As Long
Dim nbNroCuota As Byte



'==================================
Private Sub cmdArchivCentro_Click()
'==================================
'Crea el archivo R521.txt para ser enviado al circulo
' Ver la estructura mas adelante

Dim sCadena As String

'COMIENZA
Screen.MousePointer = vbHourglass
PB.Visible = True

Mensaje24 "Abriendo DTO.dbf...."
If adoM.State = adStateOpen Then adoM.Close
sCadena = "SELECT * FROM DTO;"
adoM.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText

'ATENCION la carpeta es CP y no /ARCHIVOS DE PROGRAMA/CP
Mensaje24 "Abriendo R521.txt..."
Open "C:\CP\R521.txt" For Output As 1


'Print #1, Label1.Caption & vbCrLf & "No Cobro", " Valor"

'recorre DTO.dbf
PB.Min = 0
PB.Max = adoM.RecordCount

adoM.MoveFirst
Dim sMome As String
Dim nMome As Integer
Dim sMome1 As String
Do While Not adoM.EOF
            If adoM.AbsolutePosition Mod 100 = 0 Then Mensaje24 "Recorriendo Registros   " & CStr(adoM.AbsolutePosition)
            PB.Value = adoM.AbsolutePosition
            'Print #1, Format(lSocio, "0000000"), Format(tM1, "#,#0.00")
            'Print #1, "Total: ", Format(sTot, "#,#0.00")
            If IsNull(adoM!COP) Then
                nMome = 0
            Else
                nMome = adoM!COP
            End If
            sMome1 = Format(adoM!importe, "00000000.00")
            'xxx (3) 510 jubilado 520 pensionista
            'xxxxxxx (7) Numero pasivo
            'xx (2) Coparticipe COP , si es jubilado es cero
            'xxx (3) rubro 521 es cpr
            'Importe 8 + 2 (8 enteros 2 decimales)
            ' (4) mmaa mes y año del descuento
            ' (11) documento
            '(12) Numero de solicitud
            sMome = Format(adoM!clase, "000") & Format(adoM!Numero, "0000000") & _
                            Format(nMome, "00") & Format(adoM!rubro, "000") & _
                            Left(sMome1, 8) & Right(sMome1, 2) & Format(adoM!actual, "0000") & _
                            Format(adoM!cedula, "00000000000") & "000000000000"
            Print #1, sMome
            adoM.MoveNext
Loop

DoEvents
Screen.MousePointer = vbDefault
PB.Visible = False

Mensaje24 "Cerrando DTO.dbf...."
adoM.Close
Set adoM = Nothing


Close #1
Mensaje24 ""

End Sub



'==================================
Private Sub cmdArchivoJefat_Click()
'==================================
  
Dim sCadena As String

'COMIENZA
Screen.MousePointer = vbHourglass
PB.Visible = True

Mensaje24 "Abriendo Jefatura.dbf...."
If adoM.State = adStateOpen Then adoM.Close
sCadena = "SELECT * FROM Jefatura;"
adoM.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText

Mensaje24 "Abriendo archivo Raaaamm.txt..."
Open "C:\CP\R" & vptMesPresup & ".txt" For Output As 1


'Print #1, Label1.Caption & vbCrLf & "No Cobro", " Valor"

'recorre DTO.dbf
PB.Min = 0
PB.Max = adoM.RecordCount

adoM.MoveFirst
Dim sMome As String
'Dim nMome As Integer
Dim sMome1 As String
Do While Not adoM.EOF
            If adoM.AbsolutePosition Mod 100 = 0 Then Mensaje24 "Recorriendo Registros   " & CStr(adoM.AbsolutePosition)
            PB.Value = adoM.AbsolutePosition
            'Print #1, Format(lSocio, "0000000"), Format(tM1, "#,#0.00")
            'Print #1, "Total: ", Format(sTot, "#,#0.00")
            'If IsNull(adoM!Numero_Cob) Then
            '    nMome = 0
            'Else
            '    nMome = adoM!Numero_Cob
            'End If
            sMome1 = Format(adoM!Alta, "00000000.00")
            sMome = "0416" & _
                            vptMesPresup & _
                            Left$(adoM!numero_ced, 1) & _
                            Mid$(adoM!numero_ced, 3, 3) & _
                            Mid$(adoM!numero_ced, 7, 3) & _
                            Left(sMome1, 8) & _
                            Right(sMome1, 2)
             Print #1, sMome
            adoM.MoveNext
Loop

DoEvents
Screen.MousePointer = vbDefault
PB.Visible = False

Mensaje24 "Cerrando Raaaamm.txt"
adoM.Close
Set adoM = Nothing


Close #1
Mensaje24 ""

End Sub




'==================================
Private Sub cmdCentroPol_Click()
'==================================
    If Not VerificaTablaCP Then
        MsgBox "No abre tabla DTO"
        GoTo Termi1
    End If
    Mensaje25 ""
    PB.Visible = True
   Screen.MousePointer = vbHourglass

    'Dim dFechaVtoActual As Date
    Dim sM As String
    dFechaVtoActual = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
    sM = mfInvierteMes(CStr(dFechaVtoActual))

    '1 verifica si ya no se realizó
    Mensaje24 "Verificando..."
    If Not rCierraMes_VerifMesEsteCompleto(3) Then
        If MsgBox("Cancela la generación?", vbYesNo + vbQuestion, "Tarea ya realizada!!") = vbYes Then
                    GoTo Termi1
        End If
    End If
    
    '2 Rearma PrePago
    'If MsgBox("Rearma la tabla PrePagos?", vbYesNo + vbQuestion, "Disquete a Jefatura!!") = vbYes Then
    '               msRearmaPrePago
    'End If


    '3 vacia la tabla y la abre
    Mensaje24 "Vaciando y abriendo...."
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "DELETE * FROM DTO;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    DoEvents
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM DTO;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Mensaje24 "Reg. " & adoM.RecordCount
    
    DoEvents
    'MsgBox "Continua?", vbYesNo + vbQuestion, "Info"


    '4 Abre PrePagos y revisa que los valores sean correctos
    Mensaje24 "Abriendo PrePago...."
    Dim sCadena As String
    If adoClie.State = adStateOpen Then adoClie.Close
    'sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
    '    "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc WHERE " & _
    '    "(tbl_socios.CodSitLab = 3 OR tbl_socios.codSitLab = 4) AND " & _
    '    "tbl_prepago.pp_presup ='" & vptMesPresup & "' ORDER BY cint(tbl_socios.nrocob);"
    sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
        "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc " & _
        "WHERE (tbl_socios.CodSitLab = 3 OR tbl_socios.codSitLab = 4) " & _
        "AND tbl_prepago.pp_presup ='" & vptMesPresup & "' " & _
       "ORDER BY CLNG(tbl_socios.nrocob),  tbl_socios.codsitlab, tbl_socios.cop;"
   adoClie.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoClie.RecordCount < 1 Then
        MsgBox "Sin Socios"
        Exit Sub
    Else
        Mensaje25 "Cant. Registros: " & CStr(adoClie.RecordCount)
    End If
    DoEvents
    
    'vENTANA MUY BUENA para buscar errores
    'Set fjMome.DataGrid1.DataSource = adoClie
    'fjMome.Show vbModal
    'MsgBox "hola"
    
    '5 La recorre
    PB.Min = 0
    PB.Max = adoClie.RecordCount
    
     Dim lMom As Long        'numero cobro
    Dim lMom2 As Integer    'numero cop (pues no cobro puede estar repetido)
    Dim lMom3 As Long       'CodSitLab pues un retirado puede tener mismo No que un pensionista
    Dim dTot As Single
    Dim nMom As Long
    Dim tCod As String
    Dim nCodSit As Integer
    Dim sDoc As String
    Dim sfecha As String
    
    fNumerosTablaJefatura 0
    adoClie.MoveFirst
    dTot = 0
    
    lMom = adoClie!NroCob
    Debug.Print lMom
    lMom2 = 0 + adoClie!COP
    Debug.Print lMom2
    lMom3 = adoClie!CodSitLab
        
    nMom = adoClie.RecordCount
    tCod = 0 + adoClie!COP
    nCodSit = adoClie!CodSitLab
    sDoc = adoClie!ci
    sfecha = vpbMesOperac & Right(vpnAñoOperac, 2)
    Do While Not adoClie.EOF
                    If nMom Mod 100 = 0 Then Mensaje24 "Recorriendo Registros   " & CStr(nMom)
                    nMom = nMom - 1
                    Debug.Print adoClie!NroCob, adoClie!COP, adoClie!CodSitLab
                    'iguales: el NroCobro, el Cop y el CodSitLab
                    If Not (CLng(adoClie!NroCob) = lMom And CInt(0 & adoClie!COP) = lMom2 And adoClie!CodSitLab = lMom3) Then
                                    LLenaUnRegistroCentro lMom, dTot, tCod, nCodSit, sDoc, sfecha
                                    dTot = 0
                    End If
                    lMom = adoClie!NroCob
                    lMom2 = 0 & adoClie!COP
                    lMom3 = adoClie!CodSitLab
                    dTot = dTot + adoClie!pp_Valor
                    tCod = 0 & adoClie!COP
                    nCodSit = adoClie!CodSitLab
                    sDoc = adoClie!ci
            
                    PB.Value = adoClie.AbsolutePosition
                    adoClie.MoveNext
    Loop
    'el ultimo registro
    adoClie.MoveLast
    LLenaUnRegistroCentro lMom, dTot, tCod, nCodSit, sDoc, sfecha


    '4) La imprime


    '5) Pone la marca de generado  (23)
    If Not fColocaMarcaAccion(23, vptMesPresup, "Listado a Jefatura", "", "") Then
        MsgBox "Error 2435: No se completó marca de Listado Jefatura"
    End If
   Screen.MousePointer = vbDefault

Termi1:
If adoM.State = adStateOpen Then adoM.Close
If adoClie.State = adStateOpen Then adoClie.Close

Mensaje24 ""
Set adoM = Nothing
Set adoClie = Nothing
PB.Visible = False
Unload Me
Exit Sub


End Sub




'==================================
Private Sub LLenaUnRegistroCentro(sPrm As String, sPrm As Single, tPrm As String, nPrm As Integer, sDoc As String, sfec As String)
'==================================
            adoM.AddNew
            'If IsNull(adoClie!COP) Or Len(adoClie!COP) = 0 Then
            '    adoM!clase = 520
            'Else
            '    adoM!COP = 0 + adoClie!COP
            '    adoM!clase = 510
            'End If
            If nPrm = 3 Then      'Retirado     CodSitLab=3
                adoM!clase = 510
            Else                                '4 Pensionista
                adoM!COP = tPrm
                adoM!clase = 520
            End If
            adoM!Numero = sPrm
            adoM!rubro = 892
            adoM!importe = sPrm
            adoM!actual = sfec       'mes de operacion
            adoM!Digito = 8
            adoM!cedula = SacaPuntosCedula(sDoc)
            adoM.Update

End Sub



'==================================
Private Function SacaPuntosCedula(sPrm As String) As String
'==================================
Dim nM As Integer
Dim nT As Integer
Dim sM2 As String
Dim sM3 As String


nM = Len(Trim(sPrm))
For nT = 1 To nM
    sM3 = Mid(sPrm, nT, 1)
    If IsNumeric(sM3) Then
        sM2 = sM2 & sM3
    End If
Next

SacaPuntosCedula = sM2
End Function



'==================================================
Private Sub cmdExcedidos_Click()
'==================================================
'Se crea una tabla virtual con los excedidos
'Se bajan las cuotas que estan en prepago (se pagan en la orden correspondiente)
'Excepto los excedidos
Unload Me
fjExcedidos.Show
End Sub

'==================================================
Private Sub cmdJefat_Click()
'==================================================
On Error GoTo mierro7239
    Mensaje25 ""
    If Not VerificaTablaJefatura Then
        MsgBox "No abre tabla Jefatura"
        GoTo termina
    End If
    PB.Visible = True
   Screen.MousePointer = vbHourglass

    'Dim dFechaVtoActual As Date
    Dim sM As String
    dFechaVtoActual = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
    sM = mfInvierteMes(CStr(dFechaVtoActual))
    
    '0.Probando las rutinas nuevas
    'GoTo nuevas

    '1 verifica si ya no se realizó
    'Mensaje24 "Verificando..."
    'If Not rCierraMes_VerifMesEsteCompleto(2) Then
    '    If MsgBox("Cancela la generación?", vbYesNo + vbQuestion, "Tarea ya realizada!!") = vbYes Then
    '                GoTo termina
    '    End If
    'End If
    
    '2 Rearma PrePago
    If MsgBox("Rearma la tabla PrePagos?", vbYesNo + vbQuestion, "Disquete a Jefatura!!") = vbYes Then
                   CierraMes (2)    'rearma prepago
    End If


    '3 vacia la tabla y la abre
    Mensaje24 "Vaciando y abriendo...."
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "DELETE * FROM Jefatura;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    DoEvents
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM jefatura;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText



    '4 Abre PrePagos y revisa que los valores sean correctos
    Mensaje24 "Abriendo PrePago...."
    Dim sCadena As String
    If adoClie.State = adStateOpen Then adoClie.Close
    sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
        "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc WHERE " & _
        "tbl_socios.codcatsoc = 1 AND tbl_socios.codSitLab = 1 AND " & _
        "tbl_prepago.pp_presup ='" & vptMesPresup & "' ORDER BY clng(tbl_socios.nrocob);"
    adoClie.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoClie.RecordCount < 1 Then
        MsgBox "Sin Socios"
        Exit Sub
    Else
        Mensaje25 "Cant. Registros: " & CStr(adoClie.RecordCount)
    End If

    '5 La recorre
    PB.Min = 0
    PB.Max = adoClie.RecordCount
    
    Dim lMom As Long        'numero cobro
    Dim dTot As Double
    Dim nMom As Long
    Dim sMom As String
    
    fNumerosTablaJefatura 0
    adoClie.MoveFirst
    dTot = 0
    
    lMom = adoClie!NroCob
    nMom = adoClie.RecordCount
    sMom = adoClie!ci
    Do While Not adoClie.EOF
        If nMom Mod 100 = 0 Then Mensaje24 "Recorriendo Registros   " & CStr(nMom)
        nMom = nMom - 1
        If Not adoClie!NroCob = lMom Then
            adoM.AddNew
            adoM!Codigo = fNumerosTablaJefatura(1)
            adoM!numero_ced = sMom
            adoM!Numero_Cob = lMom
            adoM!Alta = Format(dTot, "#####")
            adoM.Update
            dTot = 0
        End If
        lMom = adoClie!NroCob
        sMom = adoClie!ci
        dTot = dTot + adoClie!pp_Valor
        
        PB.Value = adoClie.AbsolutePosition
        adoClie.MoveNext
    Loop
    'el ultimo registro
     adoM.AddNew
     adoM!Codigo = fNumerosTablaJefatura(1)
    adoM!numero_ced = sMom
     adoM!Numero_Cob = lMom
     adoM!Alta = dTot
     adoM.Update


    '4) La imprime


    '5) Pone la marca de generado  (22)
    If Not fColocaMarcaAccion(22, vptMesPresup, "Listado a Jefatura", "", "") Then
        MsgBox "Error 2435: No se completó marca de Listado Jefatura"
    End If
    
    
'Rutinas nuevas
nuevas:
    '6) Genera  Jef3.dbf  que tiene el campo Nombre que Jefatura.dbf no tenia
     Mensaje24 "Vaciando y abriendo Jef3...."
    If adoP.State = adStateOpen Then adoP.Close
    adoP.Open "DELETE * FROM Jef3;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    DoEvents
    If adoP.State = adStateOpen Then adoP.Close
    adoP.Open "SELECT * FROM jef3;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
     If adoClie.State = adStateOpen Then adoClie.Close
    adoClie.Open "SELECT * FROM tbl_Socios ORDER  BY NroCob;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText

    'If adoM.State = adStateOpen Then adoM.Close
    'adoM.Open "SELECT * FROM jefatura;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    'Recorre Jefatura.dbf
    adoM.MoveFirst
    Do While Not adoM.EOF
       
        adoP.AddNew
        adoP!Numero = adoM!Numero_Cob
        adoP!nombre = ""
        adoP!mValor = adoM!Alta
         'busca el nombre del socio
        adoClie.MoveFirst
        adoClie.Find ("NroCob =" & adoM!Numero_Cob)
        If adoClie.EOF Then
            adoP("Nonbre") = ""
        Else
            adoP("Nombre") = adoClie!Apellido & " " & adoClie!nombre
        End If

        adoP.Update
        adoM.MoveNext
    Loop
    
    '7) Genera la planilla Jefatura.xls con el Número de socio, el nombre de socio y el valor del descuento.
        Dim Nro As Integer
        Nro = 1
        Mensaje24 "Abriendo aplicación Excel..."
         Set nEx = CreateObject("excel.application")
        nEx.workbooks.Open ("c:\cp\Jefat.xls")
        'agrega hoja
        Set o_Hoja = nEx.Worksheets.Add
        'Pone los títulos
        nEx.workbooks("Jefat.xls").sheets(1).cells(Nro, 1) = "Número"
        nEx.workbooks("Jefat.xls").sheets(1).cells(Nro, 2) = "Nombre"
        nEx.workbooks("Jefat.xls").sheets(1).cells(Nro, 3) = "Valor"
        Nro = 2
        'Pone los datos
        adoP.MoveFirst
        Do While Not adoP.EOF
                nEx.workbooks("Jefat.xls").sheets(1).cells(Nro, 1) = adoP!Numero
                nEx.workbooks("Jefat.xls").sheets(1).cells(Nro, 2) = adoP!nombre
                nEx.workbooks("Jefat.xls").sheets(1).cells(Nro, 3) = adoP!mValor
                adoP.MoveNext
                Nro = Nro + 1
        Loop
        'Guarda la planilla y cierra
        nEx.workbooks("Jefat.xls").SaveAs "c:\cp\Jefat.xls"
        nEx.workbooks.Close
        nEx.Quit

      
    

termina:
Screen.MousePointer = vbDefault
If adoM.State = adStateOpen Then adoM.Close
If adoP.State = adStateOpen Then adoP.Close
If adoClie.State = adStateOpen Then adoClie.Close

Mensaje24 ""
Set adoM = Nothing
Set adoP = Nothing
Set adoClie = Nothing
Set nEx = Nothing
PB.Visible = False
Unload Me
Exit Sub

mierro7239:
MsgBox "Error 7239: " & Err.Description & " " & Err.Number
MsgBox sCadena
End Sub
'==================================
Private Sub cmdCentroLista_Click()
'==================================
    
End Sub

'==================================
Private Function VerificaTablaJefatura() As Boolean
'==================================
    On Error GoTo mErr504
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM jefatura;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoM.Close
    Set adoM = Nothing
    VerificaTablaJefatura = True
    Exit Function
mErr504:
    If adoM.State = adStateOpen Then adoM.Close
    Set adoM = Nothing
VerificaTablaJefatura = False
End Function



'==================================
Private Function VerificaTablaCP() As Boolean
'==================================
    On Error GoTo mErr505
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM dto;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoM.Close
    Set adoM = Nothing
    VerificaTablaCP = True
    Exit Function
mErr505:
    If adoM.State = adStateOpen Then adoM.Close
    Set adoM = Nothing
VerificaTablaCP = False
End Function

'==================================
Private Sub cmdJefLista_Click()
'==================================
'lista la deuda indicada en tbl_Jefatura
    Dim sAvso As String
        
    On Error GoTo mErr400
    
    Mensaje24 "Abriendo Socios...."
    Dim sCadena As String
    
    sAvso = "Abriendo tbl_socios"
    If adoClie.State = adStateOpen Then adoClie.Close
    sCadena = "SELECT * FROM tbl_socios WHERE " & _
        "tbl_socios.codcatsoc = 1 AND tbl_socios.codSitLab = 1 " & _
        "ORDER BY clng(nrocob);"
    adoClie.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Debug.Print adoClie.RecordCount
    
    sAvso = "Abriendo Jefatura"
    If adoM.State = adStateOpen Then adoM.Close
    sCadena = "SELECT * FROM jefatura ORDER BY Numero_cob;"
    adoM.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Debug.Print adoM.RecordCount
        
    sAvso = "Creando tabla virtual"
    Set adoP.ActiveConnection = adoConn
    Set adoP = New ADODB.Recordset
    adoP.Fields.Append "NCobro", adInteger, 2
    adoP.Fields.Append "Nombre", adChar, 50
    adoP.Fields.Append "Valor", adSingle, 30
    
    adoP.CursorType = adOpenDynamic
    adoP.LockType = adLockOptimistic
    adoP.Open
   
   sAvso = "Recorriendo Jefatura"
    adoClie.MoveFirst
    Do While Not adoClie.EOF
        adoP.AddNew
        adoP("NCobro") = adoClie!NroCob
        adoP("Nombre") = adoClie!Apellido & "  " & adoClie!nombre
        adoM.MoveFirst
        adoM.Find ("Numero_cob =" & CLng(adoClie!NroCob))
        
        If adoM.EOF Then
            adoP("Valor") = 0
        Else
            adoP("Valor") = adoM!Alta
        End If
        adoP.Update
        adoClie.MoveNext
    Loop
  sAvso = "Imprimiendo"
  drJefatura.Caption = "Descuentos  " & vpbMesOperac & "/" & vpnAñoOperac
  drJefatura.Title = "Descuentos  " & vpbMesOperac & "/" & vpnAñoOperac

  Set drJefatura.DataSource = adoP
  drJefatura.DataMember = ""

  drJefatura.Sections(3).Controls(1).DataMember = ""
  drJefatura.Sections(3).Controls(1).DataField = "nCobro"
  drJefatura.Sections(3).Controls(2).DataMember = ""
  drJefatura.Sections(3).Controls(2).DataField = "nombre"
  drJefatura.Sections(3).Controls(3).DataMember = ""
  drJefatura.Sections(3).Controls(3).DataField = "Valor"
  'totales
  drJefatura.Sections(5).Controls(1).DataMember = ""
  drJefatura.Sections(5).Controls(1).DataField = "valor"
  
  drJefatura.Refresh
  drJefatura.Show
       
    'Set fjMome.DataGrid1.DataSource = adoP
    'fjMome.Show
    'adoP.Close
    adoM.Close
    adoClie.Close
    'Set adoP = Nothing
    Set adoM = Nothing
    Set adoClie = Nothing
    Exit Sub
    
mErr400:
    MsgBox "Error 4738: " & sAvso & "  " & Err.Description & " No." & Err.Number
End Sub
'==================================
Private Sub cmdJefListaAlf_Click()
'==================================
'lista la deuda indicada en tbl_Jefatura
'pero ordenada por nobre,
'12/02/2007
    Dim sAvso As String
    
    
    On Error GoTo mErr400
    
    Mensaje24 "Abriendo Socios...."
    Dim sCadena As String
    
    sAvso = "Abriendo tbl_socios"
    If adoClie.State = adStateOpen Then adoClie.Close
    sCadena = "SELECT * FROM tbl_socios WHERE " & _
        "tbl_socios.codcatsoc = 1 AND tbl_socios.codSitLab = 1 " & _
        "ORDER BY clng(nrocob);"
    adoClie.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Debug.Print adoClie.RecordCount
    
    sAvso = "Abriendo Jefatura"
    If adoM.State = adStateOpen Then adoM.Close
    sCadena = "SELECT * FROM jefatura ORDER BY Numero_cob;"
    adoM.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Debug.Print adoM.RecordCount
        
    sAvso = "Creando tabla virtual"
    Set adoP.ActiveConnection = adoConn
    Set adoP = New ADODB.Recordset
    adoP.Fields.Append "NCobro", adInteger, 2
    adoP.Fields.Append "Nombre", adChar, 50
    adoP.Fields.Append "Valor", adSingle, 30
    
    adoP.CursorType = adOpenDynamic
    adoP.LockType = adLockOptimistic
    adoP.Open
   
   sAvso = "Recorriendo Jefatura"
    adoClie.MoveFirst
    Do While Not adoClie.EOF
        adoP.AddNew
        adoP("NCobro") = adoClie!NroCob
        adoP("Nombre") = adoClie!Apellido & "  " & adoClie!nombre
        adoM.MoveFirst
        adoM.Find ("Numero_cob =" & CLng(adoClie!NroCob))
        
        If adoM.EOF Then
            adoP("Valor") = 0
        Else
            adoP("Valor") = adoM!Alta
        End If
        adoP.Update
        adoClie.MoveNext
    Loop
  'Ordena por nombre.
  adoP.Sort = "Nombre"
  sAvso = "Imprimiendo"
  drJefatura.Caption = "Descuentos  " & vpbMesOperac & "/" & vpnAñoOperac
  drJefatura.Title = "Descuentos  " & vpbMesOperac & "/" & vpnAñoOperac

  Set drJefatura.DataSource = adoP
  drJefatura.DataMember = ""

  drJefatura.Sections(3).Controls(1).DataMember = ""
  drJefatura.Sections(3).Controls(1).DataField = "nCobro"
  drJefatura.Sections(3).Controls(2).DataMember = ""
  drJefatura.Sections(3).Controls(2).DataField = "nombre"
  drJefatura.Sections(3).Controls(3).DataMember = ""
  drJefatura.Sections(3).Controls(3).DataField = "Valor"
  'totales
  drJefatura.Sections(5).Controls(1).DataMember = ""
  drJefatura.Sections(5).Controls(1).DataField = "valor"
  
  drJefatura.Refresh
  drJefatura.Show
       
    'Set fjMome.DataGrid1.DataSource = adoP
    'fjMome.Show
    'adoP.Close
    adoM.Close
    adoClie.Close
    'Set adoP = Nothing
    Set adoM = Nothing
    Set adoClie = Nothing
    Exit Sub
    
mErr400:
    MsgBox "Error 4738: " & sAvso & "  " & Err.Description & " No." & Err.Number
End Sub


'==================================
Private Sub cmdListaCentro_Click()
'==================================
    'lista la deuda indicada en tbl_DTO
    Mensaje24 "Abriendo Socios...."
    Dim sCadena As String
    
    
    If adoM.State = adStateOpen Then adoM.Close
    sCadena = "SELECT * FROM DTO;"
    adoM.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    Debug.Print adoM.RecordCount
        
  drCentro.Caption = "Descuentos Centro Policial " & vpbMesOperac & "/" & vpnAñoOperac
  drCentro.Title = "Descuentos  Centro Policial" & vpbMesOperac & "/" & vpnAñoOperac

  Set drCentro.DataSource = adoM
  drCentro.DataMember = ""

  drCentro.Sections(3).Controls(1).DataMember = ""
  drCentro.Sections(3).Controls(1).DataField = "clase"
  drCentro.Sections(3).Controls(2).DataMember = ""
  drCentro.Sections(3).Controls(2).DataField = "numero"
  drCentro.Sections(3).Controls(3).DataMember = ""
  drCentro.Sections(3).Controls(3).DataField = "cop"
  drCentro.Sections(3).Controls(4).DataMember = ""
  drCentro.Sections(3).Controls(4).DataField = "rubro"
  drCentro.Sections(3).Controls(5).DataMember = ""
  drCentro.Sections(3).Controls(5).DataField = "importe"
  drCentro.Sections(3).Controls(6).DataMember = ""
  drCentro.Sections(3).Controls(6).DataField = "actual"
  drCentro.Sections(3).Controls(7).DataMember = ""
  drCentro.Sections(3).Controls(7).DataField = "digito"
  drCentro.Sections(3).Controls(8).DataMember = ""
  drCentro.Sections(3).Controls(8).DataField = "cedula"
  'totales
  drCentro.Sections(5).Controls(1).DataMember = ""
  drCentro.Sections(5).Controls(1).DataField = "importe"
  
  drCentro.Refresh
  drCentro.Show
       
    'Set fjMome.DataGrid1.DataSource = adoP
    'fjMome.Show
    'adoP.Close
    adoM.Close
    'adoClie.Close
    'Set adoP = Nothing
    Set adoM = Nothing
    Set adoClie = Nothing

End Sub

'==================================================
Private Sub cmdResumen_Click()
'==================================================
'resumen del mes
Dim s1 As Double, s2 As Double, s3 As Double, s4 As Double, s5 As Double, S6 As Double, S7 As Double, s8 As Double, s9 As Double, s10 As Double, s11 As Double, s12 As Double
Dim t1 As String, t2 As String, t3 As String
   Screen.MousePointer = vbHourglass

'solo para conectar un datasource
adoClie.Open "SELECT * from tbl_Parametros;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
drCierreMesResumen.Caption = "Resumen de Cierre para " & vpbMesOperac & "/" & vpnAñoOperac
drCierreMesResumen.Title = "Resumen de Cierre para " & vpbMesOperac & "/" & vpnAñoOperac

'ordenes
Dim sMome As String
sMome = "SELECT SUM(pp_valor) as a1 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_NroOrden > 3 AND pp_mon = 'P';"
adoM.Open sMome, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a1) Then
    s1 = 0
Else
    s1 = adoM!a1
End If
t1 = "Ordenes en Pesos"
t2 = Format(s1, "#,#0.00")
t3 = Format(0#, "#,#0.00")
'DOLARES
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT SUM(pp_valor) as a2, sum(pp_valorme) as a3 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_NroOrden > 3 AND pp_mon = 'D';", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a2) Then
    s2 = 0
Else
s2 = adoM!a2
End If
If IsNull(adoM!a3) Then
    s3 = 0
Else
s3 = adoM!a3
End If
If Not s2 = 0 Then
    t1 = t1 & vbCrLf & "Ordenes en Dólares"
    t2 = t2 & vbCrLf & Format(s2, "#,#0.00")
    t3 = t3 & vbCrLf & Format(s3, "#,#0.00")
End If
'REALES
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT SUM(pp_valor) as a2, sum(pp_valorme) as a3 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_presup > 3 AND pp_mon = 'R';", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a2) Then
    s4 = 0
Else
s4 = adoM!a2
End If
If IsNull(adoM!a3) Then
    s5 = 0
Else
s5 = adoM!a3
End If
If Not s4 = 0 Then
    t1 = t1 & vbCrLf & "Ordenes en Reales"
    t2 = t2 & vbCrLf & Format(s4, "#,#0.00")
    t3 = t3 & vbCrLf & Format(s5, "#,#0.00")
End If
'OTRAS
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT SUM(pp_valor) as a2, sum(pp_valorme) as a3 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_nroorden > 3 AND NOT(pp_mon = 'P' OR pp_mon = 'D' OR pp_mon = 'R');", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a2) Then
    S6 = 0
Else
S6 = adoM!a2
End If
If IsNull(adoM!a3) Then
    S7 = 0
Else

S7 = adoM!a3
End If
If Not S6 = 0 Then
    t1 = t1 & vbCrLf & "Ordenes en Otras Monedas Ext."
    t2 = t2 & vbCrLf & Format(S6, "#,#0.00")
    t3 = t3 & vbCrLf & Format(S7, "#,#0.00")
End If

'Cuotas sociales
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT SUM(pp_valor) as a2 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_NroOrden =1;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a2) Then
    s8 = 0
Else
s8 = adoM!a2
End If
If Not s8 = 0 Then
    t1 = t1 & vbCrLf & "Cuotas sociales"
    t2 = t2 & vbCrLf & Format(s8, "#,#0.00")
    t3 = t3 & vbCrLf & Format(0#, "#,#0.00")
End If
'Ayuda social
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT SUM(pp_valor) as a2 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_NroOrden =2;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a2) Then
    s9 = 0
Else
s9 = adoM!a2
End If
If Not s9 = 0 Then
    t1 = t1 & vbCrLf & "Ayuda social"
    t2 = t2 & vbCrLf & Format(s9, "#,#0.00")
    t3 = t3 & vbCrLf & Format(0#, "#,#0.00")
End If
'Recargo
If adoM.State = adStateOpen Then adoM.Close
adoM.Open "SELECT SUM(pp_valor) as a2 fROM TBL_PrePago WHERE pp_presup ='" & vptMesPresup & "' AND pp_NroOrden =3;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
If IsNull(adoM!a2) Then
    s10 = 0
Else
s10 = 0 + adoM!a2
End If
If Not s10 = 0 Then
    t1 = t1 & vbCrLf & "Recargos"
    t2 = t2 & vbCrLf & Format(s10, "#,#0.00")
    t3 = t3 & vbCrLf & Format(0#, "#,#0.00")
End If
s11 = s1 + s2 + s4 + S6 + s8 + s9 + s10
s12 = s3 + s5 + S7
t1 = t1 & vbCrLf & "" & vbCrLf & "Total"
t2 = t2 & vbCrLf & "--------------" & vbCrLf & Format(s11, "#,#0.00")
t3 = t3 & vbCrLf & "--------------" & vbCrLf & Format(s12, "#,#0.00")


drCierreMesResumen.Sections(1).Controls(2).Caption = t1
drCierreMesResumen.Sections(1).Controls(4).Caption = t2
drCierreMesResumen.Sections(1).Controls(5).Caption = t3
drCierreMesResumen.Sections(1).Controls(9).Caption = Date

Set drCierreMesResumen.DataSource = adoClie
drCierreMesResumen.Show
DoEvents

'=======================================
'VA A MOSTRAR TODAS LAS CUENTAS
If adoM.State = adStateOpen Then adoM.Close
If adoOrd.State = adStateOpen Then adoOrd.Close
If adoClie.State = adStateOpen Then adoClie.Close
    adoClie.Open "select * FROM tbl_Socios ORDER BY NroSoc;", adoConn, adOpenStatic, adLockReadOnly, adCmdText
    adoCom.Open "select * FROM tbl_Comercios ORDER BY codigo;", adoConn, adOpenStatic, adLockReadOnly, adCmdText

    adoOrd.Open "SELECT * FROM tbl_prepago WHERE pp_presup ='" & vptMesPresup & "' ORDER BY pp_presup;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText

    '3) crea la tabla virtual
    Set adoM.ActiveConnection = adoConn
    Set adoM = New ADODB.Recordset
    Dim nM As Integer
    Dim mNomb As String
    Dim mTipo
    Dim mTam As Long
    '3.1) con los mismos campos
    For nM = 0 To adoOrd.Fields.Count - 1
        mNomb = adoOrd.Fields(nM).Name
        mTipo = adoOrd.Fields(nM).Type
        mTam = adoOrd.Fields(nM).DefinedSize
        adoM.Fields.Append mNomb, mTipo, mTam
    Next
    '3.2) con campos nuevos
    adoM.Fields.Append "Comercio", adChar, 30
    adoM.Fields.Append "Tipo", adChar, 30
    adoM.Fields.Append "Nombre", adChar, 30
    
    adoM.CursorType = adOpenDynamic
    adoM.LockType = adLockOptimistic
    adoM.Open
    '4) LLena los campos
    If adoOrd.RecordCount < 1 Then GoTo Final
    
    adoOrd.MoveFirst
    mTam = adoOrd.Fields.Count - 1
    Do While Not adoOrd.EOF
        adoM.AddNew
        'los campos comunes
        For nM = 0 To mTam
            adoM(nM) = adoOrd(nM)
        Next
        'los campos especiales
        'tipo
        Select Case adoOrd("pp_tipo")
            Case 0
                adoM("Tipo") = "Orden"
            Case 1
                adoM("Tipo") = "Cuota"
            Case 2
               adoM("Tipo") = "Ayuda"
            Case 3
               adoM("Tipo") = "Recargo"
            Case Else
               adoM("Tipo") = "?"
        End Select
        '6.4) Coloca el Comercio
        If adoOrd!pp_NroCom = 0 Then
            adoM!comercio = ""
        Else
            adoCom.MoveFirst
            adoCom.Find "Codigo =" & adoOrd!pp_NroCom
            If Not adoCom.EOF Then
                adoM!comercio = Left(Trim(adoCom!NombCom), 30)
            Else
                adoM!comercio = ""
            End If
        End If
        '6.1) Coloca el nombre
        adoClie.MoveFirst
        adoClie.Find "NroSoc =" & adoOrd!pp_NroSoc
        If Not adoClie.EOF Then
            adoM!nombre = Left(Trim(adoClie!Apellido) & "  " & Trim(adoClie!nombre), 30)
        Else
            adoM!nombre = "Desconocido"
        End If
    adoOrd.MoveNext
    Loop
  drCierreMesRes2.Caption = "Detalle de Cierre para " & vpbMesOperac & "/" & vpnAñoOperac

  drCierreMesRes2.Title = "Detalle de Cierre para " & vpbMesOperac & "/" & vpnAñoOperac

  Set drCierreMesRes2.DataSource = adoM
  drCierreMesRes2.DataMember = ""

  drCierreMesRes2.Sections(3).Controls(2).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(2).DataField = "pp_NroSoc"
  drCierreMesRes2.Sections(3).Controls(1).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(1).DataField = "pp_NroOrden"
  drCierreMesRes2.Sections(3).Controls(3).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(3).DataField = "pp_FVto"
  drCierreMesRes2.Sections(3).Controls(4).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(4).DataField = "pp_valor"
  drCierreMesRes2.Sections(3).Controls(5).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(5).DataField = "nombre"
  drCierreMesRes2.Sections(3).Controls(6).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(6).DataField = "tipo"
  drCierreMesRes2.Sections(3).Controls(7).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(7).DataField = "comercio"
  drCierreMesRes2.Sections(3).Controls(8).DataMember = ""
  drCierreMesRes2.Sections(3).Controls(8).DataField = "pp_valorme"
  'totales
  drCierreMesRes2.Sections(5).Controls(1).DataMember = ""
  drCierreMesRes2.Sections(5).Controls(1).DataField = "pp_valor"
  drCierreMesRes2.Sections(5).Controls(2).DataMember = ""
  drCierreMesRes2.Sections(5).Controls(2).DataField = "pp_valorme"
  
  drCierreMesRes2.Refresh
  drCierreMesRes2.Show
Final:
If adoM.State = adStateOpen Then adoM.Close
If adoOrd.State = adStateOpen Then adoOrd.Close
If adoClie.State = adStateOpen Then adoClie.Close
If adoCom.State = adStateOpen Then adoCom.Close
   Screen.MousePointer = vbDefault

Set adoM = Nothing
Set adoClie = Nothing
Set adoOrd = Nothing
Set adoCom = Nothing
End Sub





'==================================================
Private Sub Command1_Click()
'==================================================
msComparaValoresDosTablas
End Sub

'==================================================
Private Sub Form_Load()
'==================================================

'esconde el progress bar
PB.Visible = False

'toma el mes de operacion y lo muestra
If Not fTomaMesOperacYParametros Then
    Unload Me
    Exit Sub
End If
fMuestraMesOperac

End Sub


    


'==================================================
'Almacena un registro en tbl_Ordenes
Private Sub GuardaRegOrden(dPrm As Date, sPrm As Single, lPrm As Long, nPrm As Byte)
'==================================================
'Fecha de vto, valor del recargo, No socio, tipo de orden
On Error GoTo ERR200A

adoOrd.AddNew
adoOrd!ord_nrosoc = lPrm
adoOrd!ORD_NroCom = 0
adoOrd!ord_NroOrden = nPrm '1=CUOTA 2=AYUDA 3= Recargo
adoOrd!ORD_DEPEND = 0
adoOrd!ord_cuota = sPrm
adoOrd!ord_FEmis = dPrm
adoOrd!ord_FVto = dPrm
adoOrd!ORD_PLAN = 1
adoOrd!ord_ctasPagas = 0
adoOrd!ord_EntCta = 0
adoOrd!ord_Recarg = 0
adoOrd!ord_Mon = "P"
adoOrd!ord_mecuota = 0
adoOrd!ORD_MEPagos = 0
adoOrd!ord_cerro = CDate("01/01/1900")
adoOrd!ord_tipo = nPrm
adoOrd!ORD_Func = vpnFuncionario
adoOrd!ORD_FDia = Format(Date, "short date")
adoOrd!ORD_FHora = vptMesPresup
adoOrd.Update
Exit Sub
ERR200A:
MsgBox "Error 200A: " & Err.Description & " " & Err.Number
End Sub




'==================================================
Private Sub fCreaRegOrden()
'==================================================
    
End Sub



'==================================================
Private Sub cmdParametros_Click()
'==================================================
fjParametros.Show
End Sub






'==================================================
Private Sub fMuestraMesOperac()
'==================================================
lblMes.Caption = "Mes Operación: " & vpbMesOperac & " / " & vpnAñoOperac
End Sub



'==================================================
Private Sub cmdSalir_Click()
'==================================================
Unload Me
End Sub





'==================================================
Public Sub Mensaje24(sPrm As String)
'==================================================
lblDeta.Caption = sPrm
lblDeta.Refresh
End Sub

'==================================================
Public Sub Mensaje25(sPrm As String)
'==================================================
lblOtro.Caption = lblOtro.Caption & vbCrLf & sPrm
lblOtro.Refresh
End Sub




'==================================================
Private Function fNumerosTablaJefatura(nPrm As Byte) As String
'==================================================
' Devuelve el codigo para la tabla Jefatura.dbf
' formato: 50xxn el n se repite 15 veces y comienza en 1
'nPrm =0 inicializa
'nprm=1 devuelve numero
Static nNum As Integer  ' va contando el Numero que corresponde
Static nCta As Byte     ' cuenta de 1 a 15
Select Case nPrm
    Case 0
        nNum = 1
        fNumerosTablaJefatura = ""
    Case 1
        If nCta = 15 Then
            nCta = 0
            nNum = nNum + 1
        End If
        nCta = nCta + 1
        If nNum < 10 Then
            fNumerosTablaJefatura = "50  " & nNum
        Else
            fNumerosTablaJefatura = "50 " & nNum
        End If
End Select

End Function





'==================================================
Private Sub cmdCierraMes_Click()
'==================================================
    CierraMes (1)
End Sub




'==================================
Private Sub CierraMes(nPrm As Byte)
'==================================
'nprm = 1 Cierra Mes
'nPrm = 2 Rearma tbl_PrePago

    Dim sM As String
  

    On Error GoTo err3333
    If nPrm = 1 Then          'Cierra Mes
          'si el usuario tiene autorizacion
           If vpnNivelFuncionario < kNivel6 Then Exit Sub
        
           cmdCierraMes.Enabled = False
           cmdSalir.Enabled = False
           PB.Visible = True
           PB.Min = 0
           PB.Max = 100
           clsRecib.mfTomaNroRecibo
           
           '1) VERIFICA QUE EL MES ANTERIOR ESTA COMPLETADO
           '-----------------------------------------------
           Mensaje24 "6.Verificando repeticion..."
           'verifica que ya no se haya cerrado
           If Not rCierraMes_VerifMesEsteCompleto(1) Then
                        GoTo termina
           End If
           
           Mensaje24 "5.Verifica Mes Anterior..."
           'verifica que el mes anterior este cerrado
           If Not rCierraMes_VerifMesAnteriorEsteCompleto(1) Then
                 If MsgBox("Cancela el cierre ?", vbYesNo + vbQuestion, "Mes Anterior Incompleto!") = vbYes Then
                       GoTo termina
                 End If
           End If
           PB.Value = 5
           lblOtro.Caption = "Cierre Mes Anterior.........OK"
'===================================================
    Else        'Rearma la tabla prepago
            PB.Visible = True
            PB.Min = 0
            PB.Max = 100
            
            PB.Value = 5
            
            'Toma fecha Vencimiento  10/mm/aaaa
            If vpbMesOperac < 10 Then
                sM = CStr(vpnAñoOperac) & "0" & CStr(vpbMesOperac)
            Else
                 sM = CStr(vpnAñoOperac & vpbMesOperac)
            End If
            'Dim dFechaVtoActual As Date
            dFechaVtoActual = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
         
            
            '1 Borra los registros de este mes
            '-----------------------------------------------
            Dim sCadena As String
            If adoClie.State = adStateOpen Then adoClie.Close
            sCadena = "DELETE * FROM tbl_prepago " & _
                "WHERE pp_pRESUP ='" & sM & "';"
            adoClie.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
            If adoClie.State = adStateOpen Then adoClie.Close
            DoEvents
        
    End If
'=======================================================
    '2) Coloca las Ordenes abiertas en clsOrd.adoOrdenes
    '-----------------------------------------------
    Mensaje24 "4.Busca Ordenes..."
    If Not clsOrd.fBuscaOrdenesTodosSocios Then     'HUBO PROBLEMAS
        MsgBox "4554a: Problemas al Buscar Ordenes"
        Exit Sub
    End If
    PB.Value = 10
    Mensaje25 "Cantidad Ordenes Procesadas..." & clsOrd.adoOrdenes.RecordCount
    If clsOrd.adoOrdenes.RecordCount = 0 Then
        MsgBox "4554b: No tiene Ordenes"
        Exit Sub
    End If
    
    Mensaje24 "3.Prepara Recordset..."
    
    '3) Coloca por cuota en clsOrd.AdoM2
    '-----------------------------------------------
    clsOrd.msPreparaOrdenesAPagarEnAdoM2 (0)
    PB.Value = 15
    
    '4) Desvincula de la tabla
    '-----------------------------------------------
    Set adoM = clsOrd.adoM2
    DoEvents
    PB.Value = 17
    
    clsOrd.msTermina
    Set clsOrd = Nothing
    'PB.Max = 100
    PB.Value = 23
    
    If nPrm = 1 Then        'solo para nprm=CierraMes
    '5) Recorre los registros
    '-----------------------------------------------
        Debug.Print "5. Recorre los registros...."
        '5.1 Toma fecha Vencimiento  10/mm/aaaa
        dFechaVtoActual = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
        sM = mfInvierteMes(CStr(dFechaVtoActual))
    End If                  'solo para nprm = CierraMes
        
        '5.1 Abre las tablas que va a utilizar
        clsOrd.msInicia2
        clsPag.mfAbrePagos
        clsPreP.mfAbrePrePagos
                
        'adoClie.Open "SELECT * FROM TBL_Socios;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        'adoM.Open "SELECT * FROM TBL_PrePago;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        'adoOrd.Open "SELECT * FROM TBL_Ordenes;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

        Dim sRecargo As Single
        Dim sRecargME As Single
        Dim nMom As Long
        nMom = adoM.RecordCount
        DoEvents
        adoM.MoveFirst
        Mensaje24 "3.Prepara Recordset...adoM.MoveFirst"
        'MsgBox "Pulse..."
'=============================================================
        If nPrm = 2 Then            'si es nprm = RearmaPrePago
            Debug.Print "Rearma Pre Pago"
            Do While Not adoM.EOF
                 If nMom Mod 100 = 0 Then Mensaje24 "Generando Registros   " & CStr(nMom)
                  nMom = nMom - 1
                 '5.2 Si esta vencido:
                 '-----------------------------------------------
                 If Not adoM("VENCIM") > dFechaVtoActual Then
                         clsPreP.mfGuardaUnPrePago adoM("socio"), _
                               adoM("NoOrden"), _
                               adoM("valorp"), _
                               Format(Date, "short date"), _
                               adoM("Vencim"), _
                               0, _
                               sM, _
                               adoM("NAuto"), _
                               adoM("comercio"), _
                               adoM("Moned"), _
                               adoM("valorme"), _
                               Format(Time, "short time"), _
                               CStr(vpnFuncionario)
                              
            
                 End If
                 adoM.MoveNext
            Loop
    
            PB.Value = 100
            GoTo termina
        End If      ''nprm=2 rearma prepago
'=====================================================
        
        Dim tAnoMes As String
        If vpbMesOperac > 9 Then
            tAnoMes = CStr(vpnAñoOperac & vpbMesOperac)
        Else
           tAnoMes = CStr(vpnAñoOperac) & "0" & CStr(vpbMesOperac)
         End If
        Do While Not adoM.EOF
            If nMom Mod 100 = 0 Then Mensaje24 "Generando Registros   " & CStr(nMom)
             nMom = nMom - 1
            '5.2 Si esta vencido:
            '-----------------------------------------------
            If adoM("VENCIM") < dFechaVtoActual Then
                'Si es cuota o recargo no coloca nada
                If adoM("NoOrden") = 1 Or adoM("NoOrden") = 2 Then
                'c:Agrega en tbl_PrePagos (la cuota/ayuda)
                  clsPreP.mfGuardaUnPrePago adoM("socio"), _
                          adoM("NoOrden"), _
                          adoM("valorp"), _
                          Format(Date, "short date"), _
                          adoM("Vencim"), _
                          0, _
                          tAnoMes, _
                          adoM("NAuto"), _
                          adoM("comercio"), _
                          adoM("Moned"), _
                          adoM("valorme"), _
                          Format(Time, "short time"), _
                          CStr(vpnFuncionario)
                         
       
                Else        'NO ES CUOTA NI AYUDA, ES ORDEN ATRASADA
                               
                   'a:Agrega el recargo en tbl_Ordenes
                   If adoM("Moned") = "" Or adoM("Moned") = "P" Then
                       sRecargo = adoM("ValorP") * vpsPorcRecarg / 100
                       sRecargME = 0
                   Else
                       sRecargo = adoM("ValorME") * vpsPorcRecarg / 100
                       sRecargME = sRecargo
                   End If
                   If clsOrd.fBuscaUnaOrden2(adoM("NoOrden")) Then
                         If clsOrd.mfAgregaRecargoEnOrdenes(sRecargo) Then
                         
                         End If
                   End If
                   'b:Agrega el recargo en tbl_Pagos
                   clsPag.mfGuardaUnPago adoM("socio"), _
                       adoM("NoOrden"), _
                       sRecargo, _
                       Format(Date, "short date"), _
                       vplNroRecibo, _
                       adoM("comercio"), _
                       adoM("Moned"), _
                       sRecargME, _
                       7, _
                       "Recargo Vto " & adoM("Vencim"), _
                       Format(Time, "short time"), _
                       CStr(vpnFuncionario)
                   vplNroRecibo = vplNroRecibo + 1
                     
                   'c:Agrega en tbl_PrePagos (la orden y el recargo)
                  clsPreP.mfGuardaUnPrePago adoM("socio"), _
                          adoM("NoOrden"), _
                          adoM("valorp"), _
                          Format(Date, "short date"), _
                          adoM("Vencim"), _
                          0, _
                          tAnoMes, _
                          adoM("NAuto"), _
                          adoM("comercio"), _
                          adoM("Moned"), _
                          adoM("valorme"), _
                          Format(Time, "short time"), _
                          CStr(vpnFuncionario)
                         
                  clsPreP.mfGuardaUnPrePago adoM("socio"), _
                          adoM("NoOrden"), _
                          sRecargo, _
                          Format(Date, "short date"), _
                          adoM("Vencim"), _
                          3, _
                          tAnoMes, _
                          adoM("NAuto"), _
                          adoM("comercio"), _
                          adoM("Moned"), _
                          sRecargME, _
                          Format(Time, "short time"), _
                          CStr(vpnFuncionario)
                End If
            '5.3 Si es del ejercicio:
            '-----------------------------------------------
            ElseIf adoM("VENCIM") = dFechaVtoActual Then
                'Agrega en tbl_Prepagos (la orden)
                clsPreP.mfGuardaUnPrePago adoM("socio"), _
                       adoM("NoOrden"), _
                       adoM("valorp"), _
                       Format(Date, "short date"), _
                       adoM("Vencim"), _
                       0, _
                       tAnoMes, _
                       adoM("NAuto"), _
                       adoM("comercio"), _
                       adoM("Moned"), _
                       adoM("valorme"), _
                       Format(Time, "short time"), _
                       CStr(vpnFuncionario)

            
            '5.4 Si es de vencimiento futuro: (no esta atrasado ni es de este ejercicio)
            '-----------------------------------------------
            Else
                'Lo ignora
            End If
            adoM.MoveNext
        Loop
    clsRecib.mfGuardaNroRecibo
    '6)Genera los registros en tbl_Ordenes
    '-------------------------------------
    'a) SE GENERA UN REGISTRO POR LA CUOTA SOCIAL
    'b) SE GENERA UN REGISTRO POR LA AYUDA SOCIAL
    Mensaje24 "2.Generando cuotas...."
    adoOrd.Open "SELECT * FROM TBL_Ordenes;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    adoClie.Open "SELECT * FROM TBL_Socios;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Dim nMome As Long
    Dim sMom As Single
    
    sMom = (90 - 70) / adoClie.RecordCount 'progress bar llega a 70
    nMome = adoClie.RecordCount
    adoClie.MoveFirst
    Do While Not adoClie.EOF
        If nMome Mod 100 = 0 Then Mensaje24 "Generando Cuotas  " & CStr(nMome)
        nMome = nMome - 1
        'guarda la cuota social
        Select Case adoClie!CodCatSoc
            Case 1          'activo
                If Not vpsCuotaSAct = 0 Then
                GuardaRegOrden dFechaVtoActual, _
                        vpsCuotaSAct, adoClie!NroSoc, 1
                 End If
                        
            Case 2          'honorario
                If Not vpsCuotaSHon = 0 Then
                GuardaRegOrden dFechaVtoActual, _
                        vpsCuotaSHon, adoClie!NroSoc, 1
                End If
            Case 3          'cooperador
                If Not vpsCuotaSCop = 0 Then
                GuardaRegOrden dFechaVtoActual, _
                        vpsCuotaSCop, adoClie!NroSoc, 1
                End If
           Case Else
        End Select
        'guarda la ayuda social
        If adoClie!ayuda And Not vpsAyuda = 0 Then
            GuardaRegOrden dFechaVtoActual, _
                        vpsAyuda, adoClie!NroSoc, 2
        End If
        PB.Value = CInt(adoClie.AbsolutePosition * sMom + 70)
         adoClie.MoveNext
    Loop
    Mensaje25 "Se generaron las cuotas..."
    
    adoClie.Close
    adoOrd.Close
    adoM.Close
    
    '7) SE COPIAN LOS REGISTROS RECIEN CREADOS
    ' EN LA TBL_PREPAGO
    '----------------------------------------
    Dim bMome As Boolean
    Dim sMom1 As String
    Mensaje24 "1.Inserta Registros en PrePagos...."
    sM = mfInvierteMes(CStr(dFechaVtoActual))
    If adoOrd.State = adStateOpen Then adoOrd.Close
    sMom1 = "SELECT * FROM tbl_ordenes " & _
        "WHERE ord_NroOrden < 3 AND " & _
        "ord_FVto =#" & sM & "#;"
    If adoOrd.State = adStateOpen Then adoOrd.Close
    adoOrd.Open sMom1, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
     PB.Value = 95
    If adoOrd.RecordCount > 0 Then
        adoOrd.MoveFirst
        Do While Not adoOrd.EOF
            bMome = clsPreP.mfGuardaUnPrePago(adoOrd("ord_nrosoc"), _
                       adoOrd("ord_nroorden"), _
                       adoOrd("ord_cuota"), _
                       Format(Date, "short date"), _
                       adoOrd("ord_fvto"), _
                       adoOrd("ord_tipo"), _
                       tAnoMes, _
                       adoOrd("ord_auto"), _
                       adoOrd("ord_nrocom"), _
                       adoOrd("ord_mon"), _
                       0, _
                       Format(Time, "short time"), _
                       CStr(vpnFuncionario))
                       
            adoOrd.MoveNext
        Loop
    End If
    
    '8) SE COLOCA LA MARCA DE COMPLETADO
    If Not fColocaMarcaAccion(21, vptMesPresup, "Completado Prepago", "", "") Then
        MsgBox "Error 2433: No se completó marca de PrePago"
    End If
    PB.Value = 100

termina:
    If adoM.State = adStateOpen Then adoM.Close
    If adoP.State = adStateOpen Then adoP.Close
    If adoOrd.State = adStateOpen Then adoOrd.Close
    If adoClie.State = adStateOpen Then adoClie.Close
    
    Mensaje24 ""
    Set adoP = Nothing
    Set adoM = Nothing
    Set adoOrd = Nothing
    Set adoClie = Nothing
    cmdSalir.Enabled = True
    PB.Visible = False
    Set clsOrd = Nothing
    Set clsPag = Nothing
    Set clsPreP = Nothing
    Exit Sub

err3333:
MsgBox "Error 3333: " & Err.Description & " " & Err.Number & " pValor= " & PB.Value
End Sub



'==================================================
Private Sub msRearmaPrePago()
'==================================================
 'esta rutina NO LA UTILIZO MAS
 'esta incluida en CiarraMes(2)
 'La dejo un tiempo solo por las dudas
    'On Error GoTo err3335
    
    PB.Visible = True
    PB.Min = 0
    PB.Max = 100
    
    PB.Value = 5
    
    'Toma fecha Vencimiento  10/mm/aaaa
    Dim sM As String
    If vpbMesOperac < 9 Then
        sM = CStr(vpnAñoOperac) & "0" & CStr(vpbMesOperac)
    Else
         sM = CStr(vpnAñoOperac & vpbMesOperac)
    End If
    'Dim dFechaVtoActual As Date
    dFechaVtoActual = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
 
    
    '1 Borra los registros de este mes
    '-----------------------------------------------
    Dim sCadena As String
    If adoClie.State = adStateOpen Then adoClie.Close
    sCadena = "DELETE * FROM tbl_prepago " & _
        "WHERE pp_pRESUP ='" & sM & "';"
    adoClie.Open sCadena, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoClie.State = adStateOpen Then adoClie.Close
    DoEvents

    '2) Coloca las Ordenes abiertas en clsOrd.adoOrdenes
    '-----------------------------------------------
    Mensaje24 "4.Busca Ordenes..."
    If Not clsOrd.fBuscaOrdenesTodosSocios Then     'HUBO PROBLEMAS
        MsgBox "4554a: Problemas al Buscar Ordenes"
        Exit Sub
    End If
    PB.Value = 10
    Mensaje25 "Cantidad Ordenes Procesadas..." & clsOrd.adoOrdenes.RecordCount
    If clsOrd.adoOrdenes.RecordCount = 0 Then
        MsgBox "4554b: No tiene Ordenes"
        Exit Sub
    End If
    
    Mensaje24 "3.Prepara Recordset..."
    
    '3) Coloca por cuota en clsOrd.AdoM2
    '-----------------------------------------------
    clsOrd.msPreparaOrdenesAPagarEnAdoM2 (0)
    PB.Value = 15
    
    '4) Desvincula de la tabla
    '-----------------------------------------------
    Set adoM = clsOrd.adoM2
    DoEvents
    PB.Value = 17
    
    clsOrd.msTermina
    Set clsOrd = Nothing
    'PB.Max = 100
    PB.Value = 23

    '5) Recorre los registros
    '-----------------------------------------------
        
        '5.1 Abre las tablas que va a utilizar
        clsOrd.msInicia2
        clsPag.mfAbrePagos
        clsPreP.mfAbrePrePagos
                

        Dim sRecargo As Single
        Dim sRecargME As Single
        Dim nMom As Long
        nMom = adoM.RecordCount

        adoM.MoveFirst
        Do While Not adoM.EOF
            If nMom Mod 100 = 0 Then Mensaje24 "Generando Registros   " & CStr(nMom)
             nMom = nMom - 1
            '5.2 Si esta vencido:
            '-----------------------------------------------
            If Not adoM("VENCIM") > dFechaVtoActual Then
                    clsPreP.mfGuardaUnPrePago adoM("socio"), _
                          adoM("NoOrden"), _
                          adoM("valorp"), _
                          Format(Date, "short date"), _
                          adoM("Vencim"), _
                          0, _
                          sM, _
                          adoM("NAuto"), _
                          adoM("comercio"), _
                          adoM("Moned"), _
                          adoM("valorme"), _
                          Format(Time, "short time"), _
                          CStr(vpnFuncionario)
                         
       
            End If
            adoM.MoveNext
        Loop

    PB.Value = 100

termina:
    If adoM.State = adStateOpen Then adoM.Close
    If adoP.State = adStateOpen Then adoP.Close
    If adoOrd.State = adStateOpen Then adoOrd.Close
    If adoClie.State = adStateOpen Then adoClie.Close
    
    Mensaje24 ""
    Set adoP = Nothing
    Set adoM = Nothing
    Set adoOrd = Nothing
    Set adoClie = Nothing
    PB.Visible = False
    Set clsOrd = Nothing
    Set clsPag = Nothing
    Set clsPreP = Nothing
    Exit Sub

err3335:
MsgBox "Error 333a: " & Err.Description & " " & Err.Number & " pValor= " & PB.Value
End Sub



'==================================
Private Sub msComparaValoresDosTablas()
 '==================================
   
    If adoM.State = adStateOpen Then adoM.Close
    adoM.Open "SELECT * FROM TBL_PrePago ORDER BY pp_NroSoc;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    'If adoP.State = adStateOpen Then adoP.Close
    'adoP.Open "SELECT * FROM TBL_PrePagoOld ORDER BY pp_NroSoc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    Dim lMom As Long
    Dim dTot As Double
    Dim sMome As String
            
    adoM.MoveFirst
    lMom = adoM!pp_NroSoc
    dTot = adoM!pp_Valor
    Do While Not adoM.EOF
        If Not adoM!pp_NroSoc = lMom Then
            Mensaje24 CStr(lMom)
            If adoP.State = adStateOpen Then adoP.Close
            sMome = "SELECT SUM(pp_valor) as a1 fROM TBL_PrePagoOld WHERE pp_presup ='" & vptMesPresup & "' AND pp_NroSoc =" & lMom & ";"
            adoP.Open sMome, adoConn, adOpenKeyset, adLockOptimistic, adCmdText
            If Not Round(adoP!a1, 0) = Round(dTot, 0) Then
                Debug.Print lMom & "  " & adoP!a1 & "   " & dTot
            End If
            dTot = 0
        End If
        lMom = adoM!pp_NroSoc
        dTot = dTot + adoM!pp_Valor
        
        adoM.MoveNext
    Loop
    
    adoM.Close
    adoP.Close
    Set adoM = Nothing
    Set adoP = Nothing
End Sub





