VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fjExcedidos 
   Caption         =   "Excedidos"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   Icon            =   "fjExcedidos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLista 
      Caption         =   "Listados"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdCirc 
      BackColor       =   &H8000000A&
      Caption         =   "Círculo"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdJefa 
      BackColor       =   &H8000000A&
      Caption         =   "Jefatura"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdSigue 
      Caption         =   "Sigue"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   $"fjExcedidos.frx":628A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "fjExcedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'JEFATURA O CIRCULO: cmdCirc/cmdJefa > ms1Einicio
' 1) ms1Einicio (rCierraMes_VerifMesEsteCompleto(31 y 32))
' 2) ms1Einicio (mf1Einicio_1AbrePlanilla)(c:\cp\CPExcedidos.xls)
' 3) ms1Einicio (mf1Einicio_2SumaExcedidos) a) Crea el adoExced y llena los campos: lCobr y sVal y va sumando el total en sTot
' 4) ms1Einicio (mf1Einicio_2SumaExcedidos) b) Verifica que existen todos los socios
' 5) cmdListado llama al fjMiReporte.Show
' 6) cmdSigue: ejecuta mf2Sigue_1Parte:
'       a) Toma No de Recibo  (cRecib.mfTomaYGuardaNroRecibo)
'       b) Selecciona los registros  de la tbl_PrePago que sean los movimientos de los excedidos de este mes
'       c) Coloca en AdoAux los registros que estan para cobrar (tomados de JEFATURA o DTO)
' 7) cmdSigue: mf2sigue_2Parte:
        'Recorre el adoAux donde están los registros que se enviaron para descontar
        'a)Toma el Nùmero de cobro, segùn sea jefatura o Centro
        'b) filtra los registros del PrePago de UN No Cobro
        'c) los copia en adoaux2 (los registros del socio con No Cobro = lCobr)
        ' d) Busca en el adoExced si esta excedido
        ' e)Si no està excedido (mf2Sigue_2Parte_NoExced): descuenta todos los documentos que estan en adoPrinc (la tabla prepago)
        ' f) Si está excedido (mf2Sigue_2Parte_Exced): No descuenta nada. (No debería descontar lo que puede?
        'NO, si está excedido es que no tiene nada para cobrar, por lo tanto no se le puede descontar nada)



Option Explicit
Public n As Object

Dim cRecib As New clsRecibos
Dim cOrd As New clsOrdenes
Dim cPag As New clsPagos

Dim nPrmEx As Byte      '1=jefatura  2=CIRCULO
Dim nSigue As Byte      '1=sigue la 1a vez, 2= sigue la segunda

Dim adoPrinc As New ADODB.Recordset
Dim adoExced As New ADODB.Recordset 'tabla con los excedidos
Dim adoOrden As New ADODB.Recordset 'tabla momentanea con los registros de un excedido
Dim adoAux As New ADODB.Recordset    'auxiliar
Dim adoAux2 As New ADODB.Recordset

Dim HayError As Boolean



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub Form_Load()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        Label1.Caption = ""
        Label2.Caption = "Atención: Previamente debe guardar en c:\cp\excedidos.xls" & vbNewLine & _
            "los datos de los excedidos. (Nro Cobro y valor)." & vbNewLine & _
            "--> El último renglón debe completarse con con ceros." & vbNewLine & _
            "--> Es aconsejable hacer antes una copia de la base de datos." & vbNewLine & _
            "--> La lista de descuentos se toma de los Jefatura.dbf y dto.dbf." & vbNewLine & _
            "--> En el archivo c:\cp\MovAAAAMMx.txt queda un detalle de esta operación" & vbNewLine & _
            "(donde AAAA es el año y MM el mes que se está procesandoy x=1 es Jefatura y x=2 es Círculo)"
        Mensaje3 ""
        
        cmdSigue.Enabled = False
        cmdLista.Enabled = False
        pb.Enabled = False
        nSigue = 1
End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub cmdCirc_Click()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        nPrmEx = 2
        cmdCirc.Enabled = False
        cmdJefa.Enabled = False
        ms1Einicio

End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub cmdJefa_Click()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        nPrmEx = 1
        cmdCirc.Enabled = False
        cmdJefa.Enabled = False
        ms1Einicio
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub cmdSalir_Click()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        Unload Me
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub cmdSigue_Click()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        If nSigue = 1 Then
            mf2Sigue_1Parte
        Else
           mf2Sigue_2Parte
        End If
        
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub ms1Einicio()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        '1=jefatura
        '2=Circulo
        
        'vefica que ya no se haya realizado
        HayError = False
        Mensaje3 ("Verificando que el mes este completo...")
        If nPrmEx = 1 Then
            If Not rCierraMes_VerifMesEsteCompleto(31) Then
                    If MsgBox("Cancela el Pago?", vbYesNo + vbQuestion, "Tarea ya realizada!!") = vbYes Then
                                Exit Sub
                    End If
            End If
        Else
            If Not rCierraMes_VerifMesEsteCompleto(32) Then
                    If MsgBox("Cancela el Pago?", vbYesNo + vbQuestion, "Tarea ya realizada!!") = vbYes Then
                                Exit Sub
                    End If
            End If
        End If
        MouseOff
        
        Mensaje3 "Abriendo planilla Excedidos..."
        If Not mf1Einicio_1AbrePlanilla Then
            MouseOn
            Exit Sub
        End If
        
        Mensaje3 "Sumando Excedidos..."
        If nPrmEx = 1 Then
            Label1.Caption = "Pagos en Jefatura de " & Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
            Label2.Caption = "A continuación se ejecutarán los pagos," & vbCrLf & _
            "el total de Excedidos es de $" & Format(mf1Einicio_2SumaExcedidos, "#,#0.00")
        Else
            Label1.Caption = "Excedidos en Centro Policial de " & Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
            Label2.Caption = "A continuación se ejecutarán los pagos," & vbCrLf & _
            "el total de Excedidos es de $" & Format(mf1Einicio_2SumaExcedidos, "#,#0.00")
        End If
        
        
        Label2.Refresh
        Mensaje3 "Verificando Socios..."
        mf1Einicio_3VerificaSocios
     
        Mensaje3 "Para ejecutar los pagos, pulse SIGUE"
        If Not HayError Then
            cmdJefa.Enabled = False
            cmdCirc.Enabled = False
            cmdLista.Enabled = False
            cmdSigue.Enabled = True
        End If
        MouseOn
End Sub






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf1Einicio_1AbrePlanilla() As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        On Error GoTo mError
        
        'abre la planilla pero no la carga
        Mensaje3 "Creando aplicación Excel..."
        Set n = CreateObject("excel.application")
        Mensaje3 "Abriendo aplicación Excel..."
        n.workbooks.Open ("c:\cp\CPExcedidos.xls")
        Mensaje3 "Activando aplicación..."
        n.workbooks("CPExcedidos.xls").Activate
        mf1Einicio_1AbrePlanilla = True
        Exit Function
mError:
        MsgBox "Error 356: Al abrir la planilla CPExcedidos.xls " & vbCrLf & "Debe encontrarse en el directorio C:\CP de esta máquina" & vbCrLf & Err.Number & Err.Description
        mf1Einicio_1AbrePlanilla = False
End Function

 
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf1Einicio_2SumaExcedidos() As Single
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        ' suma excedidos y guarda resultado en
        ' cptexto.tmp
        
        Dim sTot As Single
        Dim sMome As Single
        Dim nia As Integer
        Dim lSocio As Long          'No Cobro
        Dim tM1 As String
        
        Mensaje3 "Abriendo Archivo de Salida c:\cp\MovAAAAMMx.txt..."
        Open "c:\cp\Mov" & Left(vptMesPresup, 4) & _
            Right(vptMesPresup, 2) & nPrmEx & ".txt" For Output As 1
        
        Mensaje3 "Creando ado auxiliar..."
        Set adoExced.ActiveConnection = adoConn
        Set adoExced = New ADODB.Recordset
        ' con 2 campos: No cobro y valor
        adoExced.Fields.Append "lCobr", adInteger, 2
        adoExced.Fields.Append "sVal", adSingle
        adoExced.Open
        
        Mensaje3 "Recorriendo planilla para sumar excedidos..."
        ' Graba titulos del archivo de salida: cptexto.tmp
        Print #1, Label1.Caption & vbCrLf & "No Cobro", " Valor"
        nia = 2     'SEGUNDA FILA
        ' Toma el valor
        sMome = n.workbooks("CPExcedidos.xls").sheets(nPrmEx).cells(nia, 2)
        ' Toma el No. de socio
        lSocio = n.workbooks("CPExcedidos.xls").sheets(nPrmEx).cells(nia, 1)
        sTot = sMome
        ' Hasta que no encuentre 0
        Do While Not sMome = 0
            ' Formatea el valor
            tM1 = Format(sMome, "#,#0.00")
            ' Lo guarda en el archivo de texto para seguridad
            'Print #1, Space(2 * (10 - Len(CStr(lSocio)))) & _
            '    lSocio & Space(5 + 2 * (12 - Len(tM1))) & Format(tM1, "#,#0.00")
            Print #1, Format(lSocio, "0000000"), Format(tM1, "#,#0.00")
            ' Agregar un registro nuevo
            adoExced.AddNew
            ' Guarda los datos
            adoExced!lcobr = lSocio
            adoExced!sVal = sMome
            adoExced.Update
            Debug.Print "[" & Space(2 * (10 - Len(CStr(lSocio)))) & lSocio & "]"
            nia = nia + 1
            ' Otra vez toma el valor
            sMome = n.workbooks("CPExcedidos.xls").sheets(nPrmEx).cells(nia, 2)
            ' Otra vez toma el No socio
            lSocio = n.workbooks("CPExcedidos.xls").sheets(nPrmEx).cells(nia, 1)
            'Debug.Print sMome
            ' Va sumando los valores
            sTot = sTot + sMome
        Loop
        'Label2.Caption = "Total excedidos: " & Format(sTot, "#,#0.00")
        'Label2.Refresh
        Mensaje3 "Cerrando temporarios..."
        Print #1, "Total: ", Format(sTot, "#,#0.00")
        
        mf1Einicio_2SumaExcedidos = sTot
         'Set fjMome.DataGrid1.DataSource = adoExced
        'fjMome.Show
        Mensaje3 "Cerrando Excel..."
        n.workbooks.Close
End Function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf1Einicio_3VerificaSocios() As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        Mensaje3 "Verificando Socios..."
        pb.Enabled = True
        pb.Min = 0
        pb.Max = adoExced.RecordCount
        ' Abre la tabla SOCIOS ordenada por NroCobro
        If adoAux.State = adStateOpen Then adoAux.Close
        adoAux.Open "SELECT * FROM TBL_Socios ORDER by clng(NroCob);", adoConn, adOpenForwardOnly, adLockReadOnly
        
        ' Recorre el recordset adoExced, Solo para verificar que existen todos los socios que se van a descontar
        adoExced.MoveFirst
        Do While Not adoExced.EOF
            pb.Value = adoExced.AbsolutePosition
            pb.Refresh
            ' Busca al socio
            adoAux.MoveFirst
            adoAux.Find ("NroCob =" & adoExced!lcobr)
            If adoAux.EOF Then
                MsgBox "Atención" & vbCrLf & "No encuentro en la tabla socios" & vbCrLf & "El Nro de Cobro " & _
                adoExced!lcobr & vbCrLf & "Verifique este cliente y si está mal, elimínelo de la planilla cpExcedidos.xls en la carpeta C:\CP" & _
                vbCrLf & "Esta operación se interrumpirá"
                HayError = True
            End If
            adoExced.MoveNext
        Loop
        
        pb.Value = 0
        pb.Refresh
        Mensaje3 "Cerrando tbl_socios..."
        If adoAux.State = adStateOpen Then adoAux.Close
        mf1Einicio_3VerificaSocios = True
End Function







'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub cmdLista_Click()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        ' Lista lo que se va a descontar
        MouseOff
        Mensaje3 "Cargando informe pagos..."
        drExcedidos.Caption = "Informe de Pagos "
         If nPrmEx = 1 Then 'jefatura
            drExcedidos.Title = "Informe de pagos por Jefatura " & Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
        Else               'centro
            drExcedidos.Title = "Informe de pagos por Centro " & Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
        End If
        
        Set drExcedidos.DataSource = adoPrinc
        drExcedidos.DataMember = ""
    
        drExcedidos.Sections(3).Controls(1).DataMember = ""
        drExcedidos.Sections(3).Controls(1).DataField = "NroCob"
        drExcedidos.Sections(3).Controls(2).DataMember = ""
        drExcedidos.Sections(3).Controls(2).DataField = "pp_Valor"
        drExcedidos.Sections(3).Controls(3).DataMember = ""
        drExcedidos.Sections(3).Controls(3).DataField = "pp_NroOrden"
        'totales
        drExcedidos.Sections(5).Controls(2).DataMember = ""
        drExcedidos.Sections(5).Controls(2).DataField = "pp_Valor"
        
        Mensaje3 "Actualizando informe Pagos..."
        drExcedidos.Refresh
        
        ' Lista los excedidos
        Mensaje3 "Cargando informe Excedidos..."
        drExcedidos2.Caption = "Informe de Excedidos"
        Set drExcedidos2.DataSource = adoExced
        drExcedidos2.DataMember = ""
         If nPrmEx = 1 Then 'jefatura
            drExcedidos2.Title = "Informe de Excedidos por Jefatura " & Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
        Else               'centro
            drExcedidos2.Title = "Informe de Excedidos por Centro " & Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
        End If
        
    
        drExcedidos2.Sections(3).Controls(1).DataMember = ""
        drExcedidos2.Sections(3).Controls(1).DataField = "lCobr"
        drExcedidos2.Sections(3).Controls(2).DataMember = ""
        drExcedidos2.Sections(3).Controls(2).DataField = "sVal"
        'totales
        drExcedidos2.Sections(5).Controls(2).DataMember = ""
        drExcedidos2.Sections(5).Controls(2).DataField = "sVal"
        Mensaje3 "Actualizando informe excedidos..."
        drExcedidos2.Refresh
        
        Mensaje3 ""
        MouseOn
        drExcedidos.Show
        drExcedidos2.Show
       
End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf2Sigue_1Parte() As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

        Dim sCadena As String
        Dim lNroCobro As Long
        
        MouseOff
        cmdSigue.Enabled = False
        
        'abre la tabla PAGOS
        Mensaje3 "Abriendo tabla Pagos..."
        If Not cPag.mfAbrePagos Then Exit Function
        
        
        '0) toma no. recibo
        Mensaje3 "Tomando No Recibo..."
        cRecib.mfTomaYGuardaNroRecibo
        
        '1) Toma los registros involucrados
        Mensaje3 "Abriendo registros..."
         If nPrmEx = 1 Then     'jefatura
                     sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
                        "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc WHERE " & _
                        "tbl_socios.codcatsoc = 1 AND tbl_socios.codSitLab = 1 AND " & _
                        "tbl_prepago.pp_presup ='" & vptMesPresup & "' ORDER BY clng(NroCob);"
        Else     'centro: retirados y pensionistas
                    sCadena = "SELECT * FROM tbl_prepago INNER JOIN tbl_socios " & _
                        "ON  tbl_socios.nrosoc = tbl_prepago.pp_nrosoc WHERE " & _
                        "(tbl_socios.codSitLab = 3 OR tbl_socios.codSitLab = 4) AND " & _
                        "tbl_prepago.pp_presup ='" & vptMesPresup & "' ORDER BY clng(NroCob);"
        End If
        If adoPrinc.State = adStateOpen Then adoPrinc.Close
        adoPrinc.Open sCadena, adoConn, adOpenForwardOnly, adLockReadOnly
        If adoPrinc.RecordCount < 1 Then
            MsgBox "Sin Registros"
            Exit Function
        End If
        
        ' ESCONDER ESTO !!!!!!
'        fjMome.Caption = "Registros a descontar"
'        Set fjMome.DataGrid1.DataSource = adoPrinc
'        fjMome.Show vbModal
'       If MsgBox("Si hay errores en la tabla anterior, " & vbNewLine & _
'            "detenga la ejecución de este programa." & vbNewLine & _
'            "Continua S/N", vbQuestion + vbYesNo, "Atención!!!!!") = vbNo Then
'            Unload Me
'            Exit Function
'        End If
 
        
        
        '2) Crear un adoAux solo con los numeros de cobro
'        Set adoAux.ActiveConnection = adoConn
'        Set adoAux = New ADODB.Recordset
'        adoAux.Fields.Append "lCobr", adInteger, 2
'        adoAux.Open
        
        
'        adoPrinc.MoveFirst
'        lNroCobro = adoPrinc!NroCob
'        Do While Not adoPrinc.EOF
'            If Not lNroCobro = adoPrinc!NroCob Then
'                adoAux.AddNew
'                adoAux!lcobr = lNroCobro
'                adoAux.Update
'            End If
'            lNroCobro = adoPrinc!NroCob
'            pb.Value = adoPrinc.AbsolutePosition
'            adoPrinc.MoveNext
'        Loop
'        adoAux.AddNew
'        adoAux!lcobr = lNroCobro
'        adoAux.Update
        
        
        '2 Coloca en el recordset adoAux los registros que se enviaron a cobrar
        If Not mf2Sigue_1Parte_AbreEnviados Then
            Unload Me
            Exit Function
        End If
        
        
        '3. Ofrece mostrar lo que va a descontar
        cmdLista.Enabled = True
        cmdSigue.Enabled = True
        nSigue = 2
        Label2.Caption = "Si desea detener el programa, " & vbNewLine & _
            "puede hacerlo ahora." & vbNewLine & _
            "También puede listar ahora la información que" & vbNewLine & _
            "se va a procesar, pulsando el botón LISTADOS." & vbNewLine & _
            "Para continuar presione Continuar"
        Mensaje3 ""
        MouseOn
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf2Sigue_2Parte() As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        '4 Recorre el adoAux
        '=========================================================
        ' Recorre el adoAux donde están los registros que se enviaron para descontar
        ' Busca en el adoExced si esta excedido
        ' Si no està excedido: descuenta todos los documentos que estan en adoPrinc (la tabla prepago)
        ' Si está excedido: descuenta los documentos de adoPrinc que puede
        
        MouseOff
        Label2.Caption = ""
        Dim lnCobro As Long
        Dim lnRegi As Long          'total de registros
        Dim lnMome As Long      'va contando los registros al reves
        
        Mensaje3 "Recorre el adoAux..."
        pb.Max = adoAux.RecordCount
        lnRegi = adoAux.RecordCount
        pb.Min = 0
        adoAux.MoveFirst
        Do While Not adoAux.EOF
            pb.Value = adoAux.AbsolutePosition
            lnMome = lnRegi - adoAux.AbsolutePosition
            
            'Toma el Nùmero de cobro, segùn sea jefatura o Centro
             If nPrmEx = 1 Then     'jfefatura
                lnCobro = adoAux!Numero_Cob
             Else                   'centro
                lnCobro = adoAux!Numero
             End If
            'filtra los registros del PrePago de UN No Cobro
            Label3.Caption = lnMome & " 1) "
            Label3.Refresh
            adoPrinc.Filter = "NroCob =" & lnCobro
            
            'los copia en adoaux2 (los registros del socio con No Cobro = lCobr)
            Set adoAux2 = adoPrinc
            Label3.Caption = lnMome & " 1) 2) "
            Label3.Refresh
 
            
            'Set fjMome.DataGrid1.DataSource = adoAux2
            'fjMome.Show vbModal
            'Debug.Print lnCobro
            
            'Busca si está excedido
            adoExced.MoveFirst
            adoExced.Find ("lCobr =" & lnCobro)
            
            'NO esta entre los excedidos
            If adoExced.EOF Then
                    Label3.Caption = lnMome & " 1) 2) 3)N"
                    Label3.Refresh

                    Debug.Print lnCobro & " NO"
                    Call mf2Sigue_2Parte_NoExced(lnMome)
            ' ESTA entre los excedidos
            Else
                    Label3.Caption = lnMome & " 1) 2) 3)E"
                    Label3.Refresh

                    Debug.Print lnCobro & " Excedido"
                    Call mf2Sigue_2Parte_Exced
             End If
            adoAux.MoveNext
        Loop
        pb.Value = 0
        
        
        '5) Marca como ejecutado
        
        pb.Enabled = False
        Mensaje3 ""
        cmdJefa.Enabled = True
        cmdCirc.Enabled = True
        cmdLista.Enabled = False
        cmdSigue.Enabled = False
        MouseOn
        Label3.Caption = ""
        Label3.Refresh
End Function



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf2Sigue_2Parte_NoExced(lnMome As Long) As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim sNCuota As String
Dim nTipoPago As Integer

With adoAux2

   .MoveFirst
    Do While Not .EOF
                 Debug.Print "Soc: " & !pp_NroSoc & " Ord: " & !pp_NroOrden
                 
                 '1) Actualiza la orden
                    Label3.Caption = lnMome & " 1) 2) 3)N 4)"
                    Label3.Refresh
                    nTipoPago = cOrd.mfActualizaLaOrden(!pp_NroOrden, !pp_Mon, _
                       !pp_Valor, !pp_ValorME, !pp_NroSoc, !pp_FVto)
                 'Set fjMome.DataGrid1.DataSource = cOrd.adoOrdenes
                 'fjMome.Show vbModal
                 
                 '2) Guarda en Pagos
                    Label3.Caption = lnMome & " 1) 2) 3)N 4) 5)"
                    Label3.Refresh
                    sNCuota = cOrd.adoOrdenes!ord_ctasPagas + 1 & "/" & cOrd.adoOrdenes!ORD_PLAN
                
                 If cPag.mfGuardaUnPago(!pp_NroSoc, !pp_NroOrden, _
                     !pp_Valor, Format(Date, "short date"), _
                     vplNroRecibo, _
                     !pp_NroCom, !pp_Mon, !pp_ValorME, _
                     nTipoPago, sNCuota & " " & !pp_FVto, _
                     Format(Time, "short time"), CStr(vpnFuncionario)) Then
                 End If
                 
                 '3) Guarda en el archivo MovAAAAMMx.txt
                    Label3.Caption = lnMome & " 1) 2) 3)N 4) 5) 6)"
                    Label3.Refresh
                    Print #1, "Pago", "Soc:" & !pp_NroSoc, "NoCob:" & !NroCob, "Ord:" & !pp_NroOrden, "Val:" & Format(!pp_Valor, "0.00")

            .MoveNext
    Loop
End With
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf2Sigue_2Parte_Exced() As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 
End Function






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function mf2Sigue_1Parte_AbreEnviados() As Boolean
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        ' Abre la tabla enviada a cobrar
        ' y pone los datos en el recodset adoAux
        On Error GoTo Problemas
        Mensaje3 "Creando adoAux..."
        
        Set adoAux = New ADODB.Recordset
        If nPrmEx = 1 Then
            adoAux.Open "SELECT * FROM JEFATURA;", adoConn, adOpenForwardOnly, adLockReadOnly
        Else
            adoAux.Open "SELECT * FROM DTO;", adoConn, adOpenForwardOnly, adLockReadOnly
        End If
        
        If adoAux.RecordCount < 1 Then
            MsgBox "Error 3677a: Sin registros para descontar", vbCritical, "Problemas"
            mf2Sigue_1Parte_AbreEnviados = False
            Exit Function
        End If
        mf2Sigue_1Parte_AbreEnviados = True
        Exit Function
Problemas:
        MsgBox "Error 3677b: Problemas al abrir Enviados. " & vbNewLine & _
        "Error No: " & Err.Number & vbNewLine & _
        Err.Description, vbCritical, "Problemas"
        mf2Sigue_1Parte_AbreEnviados = False
    
End Function



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub Mensaje3(sPrm As String)
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        DoEvents
        Label3.Caption = sPrm
        Label3.Refresh
End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub Form_Unload(Cancel As Integer)
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        'cierra todo
        Close #1
        n.workbooks.Close
        Set cRecib = Nothing
        Set n = Nothing
        If adoPrinc.State = adStateOpen Then adoPrinc.Close
        Set adoPrinc = Nothing
        
        If adoExced.State = adStateOpen Then adoExced.Close
        Set adoExced = Nothing
        
        If adoOrden.State = adStateOpen Then adoOrden.Close
        Set adoOrden = Nothing
        
        If adoAux.State = adStateOpen Then adoAux.Close
        Set adoAux = Nothing
        If adoAux2.State = adStateOpen Then adoAux2.Close
        Set adoAux2 = Nothing

End Sub

