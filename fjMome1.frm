VERSION 5.00
Begin VB.Form fjMome1 
   Caption         =   "Mome"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir rec. Cinta"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir recibos Tinta"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear tabla"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   $"fjMome1.frx":0000
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "fjMome1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoA As New ADODB.Recordset
Dim adoB As New ADODB.Recordset


Private Sub Command1_Click()
Mome1
End Sub

Private Sub Command2_Click()
'para Impresora: generica.
'papel: Continuo aleman estandard
'tamaño recibo: 8 cmts y un poquito
  If adoB.State = adStateOpen Then adoB.Close
       adoB.Open "SELECT * FROM tbl_RecibosEmitidos;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
Set drRecibo3.DataSource = adoB
drRecibo3.Sections(1).Controls(1).DataField = "socio"
drRecibo3.Sections(1).Controls(2).DataField = "categoria"
drRecibo3.Sections(1).Controls(3).DataField = "nombre"
drRecibo3.Sections(1).Controls(4).DataField = "totcuota"
drRecibo3.Sections(1).Controls(5).DataField = "mes"
drRecibo3.Sections(1).Controls(6).DataField = "socio"
drRecibo3.Sections(1).Controls(7).DataField = "categoria"
drRecibo3.Sections(1).Controls(8).DataField = "nombre"
drRecibo3.Sections(1).Controls(9).DataField = "totcuota"
drRecibo3.Sections(1).Controls(10).DataField = "totcuota"
drRecibo3.Sections(1).Controls(11).DataField = "mes"

drRecibo3.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Mome1()
   Screen.MousePointer = vbHourglass

    Dim dFechaVtoActual As Date
    Dim sM As String
    dFechaVtoActual = CDate(vpnPrspHst & "/" & vpbMesOperac & "/" & vpnAñoOperac)
    sM = mfInvierteMes(CStr(dFechaVtoActual))
    
    If adoB.State = adStateOpen Then adoB.Close
       adoB.Open "delete * FROM tbl_RecibosEmitidos;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

    '4 Abre PrePagos y revisa que los valores sean correctos
       Dim sCadena As String
    If adoA.State = adStateOpen Then adoA.Close
    sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
        "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc " & _
        "WHERE (tbl_socios.CodSitLab = 3 OR tbl_socios.codSitLab = 4) " & _
        "AND tbl_prepago.pp_presup ='" & vptMesPresup & "' " & _
        "ORDER BY CLNG(tbl_socios.nrocob);"
   adoA.Open sCadena, adoconn, adOpenKeyset, adLockOptimistic, adCmdText
    If adoA.RecordCount < 1 Then
        MsgBox "Sin Socios"
        Exit Sub
    End If
    DoEvents
    
    If adoB.State = adStateOpen Then adoB.Close
       adoB.Open "SELECT * FROM tbl_RecibosEmitidos;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

    '5 La recorre
 
    
    Dim lMom As Long        'numero cobro
    Dim dTot As Single
    Dim sNombre As String
    Dim sSocio As String
     
     adoA.MoveFirst
    dTot = 0
    
    lMom = adoA!NroCob
    sNombre = adoA!Apellido & " " & adoA!nombre
    sSocio = adoA!NroSoc

    Do While Not adoA.EOF
         If Not adoA!NroCob = lMom Then
            mm2 lMom, dTot, sNombre, sSocio
            dTot = 0
        End If
        lMom = adoA!NroCob
        dTot = dTot + adoA!pp_Valor
        sNombre = adoA!Apellido & " " & adoA!nombre
        sSocio = adoA!NroSoc
        adoA.MoveNext
    Loop
    'el ultimo registro
    adoA.MoveLast
            mm2 lMom, dTot, sNombre, sSocio


   Screen.MousePointer = vbDefault

termina:
If adoB.State = adStateOpen Then adoB.Close
If adoA.State = adStateOpen Then adoA.Close

Set adoB = Nothing
Set adoA = Nothing

Unload Me
Exit Sub


End Sub



Private Sub mm2(sP1 As Long, sP2 As Single, sP3 As String, sp4 As String)
            adoB.AddNew
            adoB!nombre = sP3
            adoB!socio = sp4
            adoB!categoria = sP1
            adoB!TOTCUOTA = sP2
            adoB!mes = "Dic 2002"
            adoB.Update

End Sub




Private Sub Command4_Click()
    If adoB.State = adStateOpen Then adoB.Close
    adoB.Open "SELECT * FROM tbl_RecibosEmitidos;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

    'Printer.Height = 17500
    Printer.CurrentX = 0
    Printer.CurrentY = 0
   With adoB
    
    adoB.MoveFirst
    Do While Not adoB.EOF
        Printer.Print Space(8) & mfCompleta(!socio, 6) & Space(3) & mfCompleta(!categoria, 6) & Space(25) & mfCompleta(!socio, 10) & Space(5) & mfCompleta(!categoria, 10)
        Printer.Print
        Printer.Print
        Printer.Print Space(1) & mfCompleta(!nombre, 25) & Space(20) & !nombre
        Printer.Print
        Printer.Print
        Printer.Print Space(35) & !mes & Space(23) & Format(!TOTCUOTA, "#,#0.00")
        Printer.Print
        Printer.Print
        Printer.Print Space(25) & Format(!TOTCUOTA, "#,#0.00")
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print Space(3) & !mes & Space(2) & Format(!TOTCUOTA, "#,#0.00")
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
    'dejo el 4to recibo vacio por problemas impresora.
        If .AbsolutePosition Mod 3 = 0 Then
            Printer.NewPage
        End If
         
        adoB.MoveNext
        
    Loop
    Printer.EndDoc
    End With
    adoB.Close
    Set adoB = Nothing
End Sub

Private Sub Command5_Click()
    Screen.MousePointer = vbHourglass

    If adoB.State = adStateOpen Then adoB.Close
    adoB.Open "SELECT  * FROM tbl_socios ORDER BY nrosoc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

     adoA.MoveFirst
    Do While Not adoA.EOF
        Label2.Caption = adoA!NroSoc
        adoA.MoveNext
    Loop

    Screen.MousePointer = vbDefault

    If adoB.State = adStateOpen Then adoB.Close

    Set adoB = Nothing


End Sub

