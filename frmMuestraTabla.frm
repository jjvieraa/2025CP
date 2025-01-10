VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjMuestraTabla 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMuestraTabla.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2895
   End
End
Attribute VB_Name = "fjMuestraTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'jv   &H00C0FFFF&   AMARILLO
'   &H00C0FFC0&   VERDE
   
   
        Option Explicit
        
        
        Const Amarillo = 1
        Const Verde = 2


Dim vCampoOrd As String
Dim vTipoBusqueda As Byte   '1=numerica 2=texto
Dim Funcion As String
Dim adoMT As ADODB.Recordset         'muestra tabla
Dim cD As New clsDepend

Const vkAlfabetico = 2
Const vkNumerico = 1


Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub dg1_Click()
Funcion = "ENT"
End Sub


Private Sub DG1_DblClick()
Funcion = "ENT"
End Sub

Private Sub DG1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo mErr444
If KeyCode = vbKeyReturn Then
            Me.Hide
            If ValidaDatos Then
                Call DevuelveDatos 'traslada la informacion
                cmdSalir_Click
            End If
End If
Exit Sub
mErr444:
MsgBox "Error 342: " & Err.Description & "  " & Err.Number
End Sub


'=======================================================================================================
Private Sub DG1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'=======================================================================================================
        On Error GoTo mavso2
        If Funcion = "ENT" Then
            If ValidaDatos Then
                Call DevuelveDatos 'traslada la informacion
                cmdSalir_Click
            End If
        End If
        Exit Sub

mavso2:
        MsgBox "ERROR mavso2: " & Err.Description & " " & Err.Number
End Sub






'=======================================================================================================
Private Sub DG1_KeyPress(KeyAscii As Integer)
'=======================================================================================================
Static HHH As String
On Error GoTo mavso4
If KeyAscii = 27 Then        'escape
    Me.Hide
    Unload Me
ElseIf KeyAscii = 32 Then        'ESUN ESPACIO
    Funcion = "ESP"
    HHH = ""                        'LIMPIA LA BUSQUEDA
    adoMT.MoveFirst
ElseIf KeyAscii = vbKeyPageUp Then
    Funcion = "PGU"
    adoMT.Move (-20)
    If adoMT.BOF Then
        adoMT.MoveFirst
    End If

ElseIf KeyAscii = vbKeyPageDown Then
    Funcion = "PGD"
    adoMT.Move (20)
    If adoMT.EOF Then
        adoMT.MoveLast
    End If
ElseIf KeyAscii = vbKeyUp Then      'flecha para encima
    If Not adoMT.BOF Then adoMT.Move (-1)
    
ElseIf KeyAscii = vbKeyDown Then    'flecha para abajo
    If Not adoMT.EOF Then adoMT.Move (1)

ElseIf KeyAscii = vbKeyReturn Then
    Funcion = "ENT"
    SendKeys "{TAB}"
ElseIf KeyAscii = vbKeyTab Then
    Funcion = "ENT"
Else
    Funcion = "MOV"
    HHH = HHH & Chr(KeyAscii)
    msBuscaRegistro (HHH)
End If
Exit Sub

mavso4:
    MsgBox "ERROR mavso4: " & Err.Description & " " & Err.Number
End Sub


'=======================================================================================================
Private Sub msBuscaRegistro(TTT As String)
'=======================================================================================================
    On Error GoTo ME3326
    Dim nMsg As String
    If vTipoBusqueda = vkAlfabetico Then    'busqueda texto
        nMsg = vCampoOrd & " like '" & TTT & "*'"
        'nMsg = vCampoOrd & " ='" & TTT & "*'"
    Else                          'busqueda numerica
        If Not IsNumeric(TTT) Then
            MsgBox "Pulse solo números", vbCritical, "Error"
            Exit Sub
            
        Else
            nMsg = vCampoOrd & " =" & TTT & "*"
        End If
    End If
    adoMT.Find (nMsg)
    Exit Sub
ME3326:
    MsgBox "ERROR me3326: " & Err.Description & " " & Err.Number
End Sub



'=======================================================================================================
Private Sub Form_Unload(Cancel As Integer)
'=======================================================================================================
On Error GoTo miError22022
If adoMT.State = adStateOpen Then
    adoMT.Close
End If
Set adoMT = Nothing
Set fjMuestraTabla = Nothing
Set cD = Nothing
Exit Sub

miError22022:
    MsgBox "ERROR 220d22: Cerrando formulario  " & Err.Description
End Sub





'=======================================================================================================
Private Sub cmdSalir_Click()
'=======================================================================================================
On Error GoTo mavso6
Unload Me
Exit Sub

mavso6:
    MsgBox "ERROR mavso6: " & Err.Description & " " & Err.Number
End Sub


'=======================================================================================================
Private Sub DG1_GotFocus()
'=======================================================================================================
On Error GoTo mavso1
'DG1.SelStart = 0
'DG1.SelLength = 0       'Len(DG1.Text)
Exit Sub

mavso1:
    If Not Err.Number = 7005 Then    'CONJUNTO FILAS NO DISPONIBLES
        MsgBox "ERROR mavso1: " & Err.Description & " " & Err.Number
    End If
End Sub










'=======================================================================================================
Private Sub Form_Load()
'=======================================================================================================
     Set adoMT = New ADODB.Recordset        'muestra tabla

    On Error GoTo miError22021
    
    lblInfo.Caption = "En las tablas ordenadas alfabéticamente, " & _
    "puede dar saltos pulsando las letras correspondientes. " & _
    "Para iniciar una nueva palabra pulse ESPACIO"
    
    Set adoMT = New ADODB.Recordset
    Set adoMT.ActiveConnection = adoConn
    If adoMT.State = adStateOpen Then adoMT.Close
    vpbVieneDeMuestraTabla = False
Select Case vpMuestraTabla
        Case kMstrSocAlf, kMstrSocAlf3, kMstrSocAlf5, kMstrSocAlf6, kMstrSocAlf7, kMstrSocAlf9
            vCampoOrd = "Otro"
            vTipoBusqueda = vkAlfabetico          'TEXTO
            fjMuestraTabla.Caption = "Socios"
            adoMT.Open "SELECT  Apellido & '  ' & Nombre as Otro,NroSoc FROM TBL_Socios " & _
                "ORDER BY Apellido,Nombre;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
        Case kMstrSoc1
            vCampoOrd = "NroSoc"
            vTipoBusqueda = vkNumerico   'NUMERICA
            fjMuestraTabla.Caption = "Socios"
            adoMT.Open "SELECT  NroSoc, Apellido & '  ' & Nombre as Otro FROM TBL_Socios " & _
                "ORDER BY NroSoc;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
         Case kMstrSocPorCobr1, kMstrSocPorNC2, kMstrSocPorNC3
            vCampoOrd = "NroCob"
            vTipoBusqueda = vkAlfabetico      'TEXTO
            fjMuestraTabla.Caption = "Socios"
            adoMT.Open "SELECT  NroCob,NroSoc, Apellido & '  ' & Nombre as Otro FROM TBL_Socios " & _
                "ORDER BY clng(NroCob);", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
       Case kMuestraComercios, kMstrComerc2, kMstrComerc3, kMstrComerc4
            vCampoOrd = "NombCom"
            vTipoBusqueda = vkAlfabetico   'texto
            fjMuestraTabla.Caption = "Comercios"
            adoMT.Open "SELECT  NombCom,Codigo FROM TBL_Comercios " & _
                "ORDER BY NombCom;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
      Case kMuestraDepend
            vCampoOrd = "NroDep"
            vTipoBusqueda = vkNumerico      'NUMERICA
            fjMuestraTabla.Caption = "Dependientes"
            cD.Inicia
            cD.vlNroSoc = vplNroSocio
            If cD.fBuscaDependientUnSocio Then
                cD.fOrdenaAdoPorDepend
                Set adoMT = cD.adoMome
             End If
      Case kMstrGastos
            vCampoOrd = "sRubro"
            vTipoBusqueda = vkNumerico      'NUMERICA
            fjMuestraTabla.Caption = "Rubros"
            adoMT.Open "SELECT  * FROM tbl_GastosRubros ORDER BY sRubro;", adoConn, adOpenKeyset, adLockOptimistic, adCmdText
      Case Else
    End Select
    If adoMT.State = adStateOpen Then
        If adoMT.RecordCount > 0 Then
            Set DG1.DataSource = adoMT
            MColumnas
        End If
    End If
    Exit Sub
miError22021:
    MsgBox "ERROR 220d21: Abriendo Nombres:  " & Err.Description
    Unload Me
End Sub
'=======================================================================================================
Private Sub MColumnas()
'=======================================================================================================
On Error GoTo mavso8
    Select Case vpMuestraTabla
        Case kMstrSocAlf, kMstrSocAlf5, kMstrSocAlf6, kMstrSocAlf9
            'DG1.Columns(0).DataField = adoMT!Otro
            'DG1.Columns(1).DataField = adoMT!NroSoc
            DG1.Columns(0).Caption = "Nombre"
            DG1.Columns(1).Caption = "Número"
            DG1.Columns(0).Width = 2594
            DG1.Columns(1).Width = 650
        Case kMstrSoc1
            DG1.Columns(0).Caption = "Número"
            DG1.Columns(1).Caption = "Nombre"
            DG1.Columns(0).Width = 650
            DG1.Columns(1).Width = 2594
         Case kMstrSocAlf3
            DG1.Columns(0).Caption = "Nombre"
            DG1.Columns(1).Caption = "Número"
            DG1.Columns(0).Width = 2594
            DG1.Columns(1).Width = 650
         Case kMstrSocPorCobr1, kMstrSocPorNC2, kMstrSocPorNC3
            DG1.Columns(0).Caption = "NoCobro"
            DG1.Columns(1).Caption = "NoSocio"
            DG1.Columns(2).Caption = "Nombre"
            DG1.Columns(0).Width = 650
            DG1.Columns(1).Width = 650
            DG1.Columns(2).Width = 2594
       Case kMuestraComercios, kMstrComerc2, kMstrComerc3, kMstrComerc4
            DG1.Columns(1).Caption = "Número"
            DG1.Columns(0).Caption = "Nombre"
            DG1.Columns(1).Width = 650
            DG1.Columns(0).Width = 2594
       Case kMuestraDepend
            DG1.Columns(0).Caption = "Número"
            DG1.Columns(1).Caption = "Nombre"
            DG1.Columns(0).Width = 650
            DG1.Columns(1).Width = 2594
            DG1.Columns(2).Visible = False
       Case kMstrGastos
            DG1.Columns(0).Caption = "Rubro"
            DG1.Columns(1).Caption = "Detalle"
            DG1.Columns(0).Width = 650
            DG1.Columns(1).Width = 2594
            
   Case Else
            cmdCerrar_Click
    End Select

Exit Sub

mavso8:
    MsgBox "ERROR mavso8: " & Err.Description & " " & Err.Number

End Sub


'=======================================================================================================
Private Sub DevuelveDatos()
'=======================================================================================================
    On Error GoTo mavso3
    vpbVieneDeMuestraTabla = True
      Select Case vpMuestraTabla
        Case kMstrSocAlf
            DG1.Col = 1
            fjIngresos2.Garantia.Text = DG1.Text
            fjIngresos2.Garantia.SetFocus
        Case kMstrSoc1
            DG1.Col = 0
            fjIngresos2.NroSoc.Text = DG1.Text
            fjIngresos2.NroSoc.SetFocus
        Case kMstrSocAlf5
            DG1.Col = 1
            fjIngresos2.NroSoc.Text = DG1.Text
            fjIngresos2.NroSoc.SetFocus
        Case kMstrSocAlf9, kMstrSocPorNC2
            DG1.Col = 1
            fjCobros.txtSocio.Text = DG1.Text
            fjCobros.txtSocio.SetFocus
        Case kMstrSocAlf3
            DG1.Col = 1
            fjOrdenes.txt(0).Text = DG1.Text
            fjOrdenes.txt(0).SetFocus
        Case kMstrSocPorCobr1
            DG1.Col = 1
            fjOrdenes.txt(0).Text = DG1.Text
            fjOrdenes.txt(0).SetFocus
        Case kMstrSocAlf6, kMstrSocPorNC3
            DG1.Col = 1
            fjInformes.Text1.Text = DG1.Text
            fjInformes.Text1.SetFocus
        Case kMuestraDepend
            DG1.Col = 0
            fjOrdenes.txt(2).Text = DG1.Text
            fjOrdenes.txt(2).SetFocus
        Case kMuestraComercios
            DG1.Col = 1
            fjOrdenes.txt(1).Text = DG1.Text
            fjOrdenes.txt(1).SetFocus
        Case kMstrComerc2
            DG1.Col = 1
            fjComercios.Text2.Text = DG1.Text
            fjComercios.Text2.SetFocus
        Case kMstrComerc3
            DG1.Col = 1
            fjPagoAComerc.txtComerc.Text = DG1.Text
            fjPagoAComerc.txtComerc.SetFocus
        Case kMstrComerc4
            DG1.Col = 1
            fjInformes.Text3.Text = DG1.Text
            fjInformes.Text3.SetFocus
        Case kMstrGastos
            DG1.Col = 0
            fjSalidasYEntradas.Text1(0).Text = DG1.Text
            fjSalidasYEntradas.Text1(0).SetFocus
        End Select
    SendKeys "{TAB}"
Exit Sub

mavso3:
    If Not Err.Number = 7005 Then   'FILAS NO DISPONIBLES
        MsgBox "ERROR mavso3: " & Err.Description & " " & Err.Number
    End If
End Sub
'=======================================================================================================
Private Function ValidaDatos() As Boolean
'=======================================================================================================
    'Select Case vpbMuestraTabla
    '     Case kMstrCiudades1, kMstrProveedPorNum
     '       If Not IsNumeric(DG1.Text) Then
     '           MsgBox "Error 3427: Validación"
     '           ValidaDatos = False
     '           Exit Function
     '       End If
     '   Case kMstrCiudades2
     '
     '   End Select
ValidaDatos = True
End Function



Private Sub miColor(nPrm As Byte)
If nPrm = Amarillo Then
    DG1.BackColor = &HC0FFFF
    cmdCerrar.BackColor = &HC0FFFF
Else
     DG1.BackColor = &HC0FFC0
    cmdCerrar.BackColor = &HC0FFC0
    'mColor = vbGreen        'RGB(0, 128, 0)
End If

End Sub


