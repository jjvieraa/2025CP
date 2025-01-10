VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form fjPagoOrden 
   Caption         =   "Pago Cuotas de Ordenes"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Escritorio\policia\Circulo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5625
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CtaSoc"
      Top             =   150
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adelanto Cta. Soc."
      Height          =   420
      Left            =   3165
      TabIndex        =   26
      Top             =   6255
      Width           =   1005
   End
   Begin VB.Data DtAyuda 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Escritorio\policia\Circulo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2490
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Data DtCtaSoc 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Escritorio\policia\Circulo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSMask.MaskEdBox tCreditos 
      Height          =   375
      Left            =   1170
      TabIndex        =   19
      Top             =   7680
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame4 
      Height          =   705
      Left            =   90
      TabIndex        =   16
      Top             =   6060
      Width           =   2265
      Begin VB.CommandButton btnConfirma 
         Caption         =   "Paga la Orden Seleccionada"
         Height          =   495
         Left            =   360
         MouseIcon       =   "frmPagoOrden.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   4920
      TabIndex        =   14
      Top             =   6060
      Width           =   2295
      Begin VB.CommandButton btnPagaTodo 
         Caption         =   "Paga Todas la Ordenes"
         Height          =   495
         Left            =   330
         MouseIcon       =   "frmPagoOrden.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   135
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   9510
      TabIndex        =   12
      Top             =   6060
      Width           =   2175
      Begin VB.CommandButton btnHaceEntrega 
         Caption         =   "Hace Una Entrega"
         Height          =   495
         Left            =   315
         MouseIcon       =   "frmPagoOrden.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   150
         Width           =   1545
      End
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   5580
      TabIndex        =   11
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Data DtPagos 
      Caption         =   "DtPagos"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Escritorio\policia\Circulo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pagos"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Data DtOrden 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Escritorio\policia\Circulo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenes pendientes de Pago hasta el presupuesto seleccionado"
      Height          =   4185
      Left            =   90
      TabIndex        =   8
      Top             =   1830
      Width           =   11565
      Begin VB.PictureBox DBGrid1 
         Height          =   3765
         Left            =   180
         ScaleHeight     =   3705
         ScaleWidth      =   11205
         TabIndex        =   9
         Top             =   300
         Width           =   11265
      End
   End
   Begin VB.TextBox txtAño 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1290
      Width           =   705
   End
   Begin VB.TextBox txtMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2190
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1290
      Width           =   405
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3450
      TabIndex        =   3
      Top             =   600
      Width           =   2925
   End
   Begin VB.TextBox txtCodSoc 
      Height          =   285
      Left            =   1830
      TabIndex        =   0
      Top             =   615
      Width           =   945
   End
   Begin VB.Data DtSocio 
      Caption         =   "DtSocio"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Escritorio\policia\Circulo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Clientes"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1845
   End
   Begin MSMask.MaskEdBox tCuota 
      Height          =   375
      Left            =   4140
      TabIndex        =   20
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Total 
      Height          =   375
      Left            =   9465
      TabIndex        =   22
      Top             =   7635
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox totAyuda 
      Height          =   375
      Left            =   6975
      TabIndex        =   24
      Top             =   7665
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Ayuda Soc"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5610
      TabIndex        =   25
      Top             =   7710
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8715
      TabIndex        =   23
      Top             =   7665
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cuota:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3270
      TabIndex        =   21
      Top             =   7710
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Creditos:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   18
      Top             =   7710
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   5100
      TabIndex        =   10
      Top             =   1350
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   1380
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hasta el Presupuesto?"
      Height          =   195
      Left            =   510
      TabIndex        =   4
      Top             =   1350
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del Socio"
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo de Socio"
      Height          =   195
      Left            =   525
      TabIndex        =   1
      Top             =   645
      Width           =   1170
   End
End
Attribute VB_Name = "fjPagoOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ent As Double
Dim sd As Double
Dim ValCta As Double

Sub Adelanta()
MsgBox "Procedimiento Adelanta"
'COMIENZA TRABAJO CON LA TABLA AYUDA
Dim mes As Integer
Dim Adelanto As Double
Dim X As Double

mes = Month(Date)

Do Until Adelanto <= 0
   ' Adelanto   = es el valor de la entrega
  mes = mes + 1
'si el valor de la entrega es sup al saldo  saldo = 0 (paga cuota)
 MsgBox "Valor  Adelanto antes de los IF  " & Adelanto
 If Adelanto >= ValCta Then
  
Adelanto = Adelanto - ValCta

MsgBox "Valor Adelanto mes 8   " & Adelanto
  
   DtCtaSoc.Recordset.AdelantoNew
        DtCtaSoc.Recordset("IDCliente") = txtCodSoc
        DtCtaSoc.Recordset("Mes") = mes
        DtCtaSoc.Recordset("Ano") = Year(Date)
        DtCtaSoc.Recordset("Valor") = ValCta
        DtCtaSoc.Recordset("Saldo") = 0
        DtCtaSoc.Recordset("Pago") = "S"
    DtCtaSoc.Recordset.Update
  
  
           
End If
'>>>>>>>>>>>>>>>Si lo que queda de la entrega no cubre el saldo<<<<<
If Adelanto < ValCta And Adelanto <> 0 Then
mes = mes + 1

     X = ValCta - Adelanto
      
      DtCtaSoc.Recordset.AdelantoNew
        DtCtaSoc.Recordset("IDCliente") = txtCodSoc
        DtCtaSoc.Recordset("Mes") = mes
        DtCtaSoc.Recordset("Ano") = Year(Date)
        DtCtaSoc.Recordset("Valor") = ValCta
        DtCtaSoc.Recordset("Saldo") = X
        DtCtaSoc.Recordset("Pago") = "N"
      DtCtaSoc.Recordset.Update
      
    DtCtaSoc.Refresh
    
     
               
               
               Exit Do
End If
     
Loop
End Sub

Sub EntAy()


'COMIENZA TRABAJO CON LA TABLA AYUDA
If DtAyuda.Recordset.RecordCount = 0 Then EntOrd: Exit Sub
  DtAyuda.Recordset.MoveFirst 'posiciona en el primer registro


Do While Not DtAyuda.Recordset.EOF
   ' ent ibox = es el valor de la entrega
  
'si el valor de la entrega es sup al saldo  saldo = 0 (paga cuota)
  If Ent >= DtAyuda.Recordset("saldo") Then
  
Ent = Ent - DtAyuda.Recordset("Saldo")

   
   DtAyuda.Recordset.Edit
        DtAyuda.Recordset("Saldo") = 0
    DtAyuda.Recordset.Update
  
  
           
End If
'>>>>>>>>>>>>>>>Si lo que queda de la entrega no cubre el saldo<<<<<
If Ent < DtAyuda.Recordset("Saldo") Then

     sd = DtAyuda.Recordset("Saldo") - Ent
      
      DtAyuda.Recordset.Edit
       
        DtAyuda.Recordset("Saldo") = sd
      DtAyuda.Recordset.Update
      
    DtAyuda.Refresh
    
    tot
               
               
               Exit Do
End If
    DtAyuda.Recordset.MoveNext
Loop
MsgBox "EOF DE AYUDA"
EntOrd
End Sub

Sub EntCta()

If DtCtaSoc.Recordset.RecordCount = 0 Then EntAy: Exit Sub
  DtCtaSoc.Recordset.MoveFirst 'posiciona en el primer registro

'COMIENZO DE  TRABAJO CON LA TABLA CTASOC

Do While Not DtCtaSoc.Recordset.EOF
   
  
'si el valor de la entrega es sup al saldo  saldo = 0 (paga cuota)
  
If Ent >= DtCtaSoc.Recordset("saldo") Then
 
Ent = Ent - DtCtaSoc.Recordset("Saldo")

   
   DtCtaSoc.Recordset.Edit
        DtCtaSoc.Recordset("Saldo") = 0
    DtCtaSoc.Recordset.Update
  

  
        
 End If
'>>>>>>>>>>>>>>>Si lo que queda de la entrega no cubre el saldo<<<<<
If Ent < DtCtaSoc.Recordset("Saldo") Then

     sd = DtCtaSoc.Recordset("Saldo") - Ent
      
      DtCtaSoc.Recordset.Edit
       
        
        DtCtaSoc.Recordset("Saldo") = sd
      
      DtCtaSoc.Recordset.Update
    
    DtCtaSoc.Refresh
    
    tot
               
               
               Exit Do
End If
 
    DtCtaSoc.Recordset.MoveNext

 Loop
 MsgBox "EOF DE CTA"
EntAy
End Sub

Sub EntOrd()
If DtOrden.Recordset.RecordCount = 0 Then Exit Sub: MsgBox "0 Reg en Ordenes"

  DtOrden.Recordset.MoveFirst 'posiciona en el primer registro


MsgBox "el valor de la entrega es de   " & Ent
Do While Not DtOrden.Recordset.EOF
   ' ent ibox = es el valor de la entrega
  
'si el valor de la entrega es sup al saldo  saldo = 0 (paga cuota)
  If Ent >= DtOrden.Recordset("saldo") Then
  
Ent = Ent - DtOrden.Recordset("Saldo")

   
   DtOrden.Recordset.Edit
        DtOrden.Recordset("Saldo") = 0
    DtOrden.Recordset.Update
  
 IngPago
          
        
End If
'>>>>>>>>>>>>>>>Si lo que queda de la entrega no cubre el saldo<<<<<
If Ent < DtOrden.Recordset("Saldo") Then

     sd = DtOrden.Recordset("Saldo") - Ent
      DtOrden.Recordset.Edit
        
        DtOrden.Recordset("Saldo") = sd
      DtOrden.Recordset.Update
      
    DtOrden.Refresh
    
    IngPago
    tot
               
    
               Exit Do
End If
DtOrden.Recordset.MoveNext

Loop
End Sub

Sub IngPago()
With DtPagos
    .Recordset.AdelantoNew
    .Recordset("Fecha") = Fecha
    .Recordset("Cliente") = DtOrden.Recordset("IDCliente")
    .Recordset("Comercio") = DtOrden.Recordset("comercio")
    .Recordset("Valor") = DtOrden.Recordset("Valor")
    .Recordset("Saldo") = DtOrden.Recordset("Saldo")
    .Recordset("Mes") = DtOrden.Recordset("Mes")
    .Recordset("Ano") = DtOrden.Recordset("Ano")
    .Recordset("Cta") = DtOrden.Recordset("Cuota")
    .Recordset("NroOrd") = DtOrden.Recordset("Ord_Nro")
    .Recordset("Int") = DtOrden.Recordset("Int")
    .Recordset("Situacion") = DtOrden.Recordset("Situacion")
    .Recordset("Rec") = 0
    .Recordset("RegNo") = 0
    .Recordset.Update
        
End With
End Sub


Sub tot()
Dim Tord As Single, tcuot As Single, Tayuda As Single
Dim s(5) As Single
'Comienza a sumar Ordenes
If DtOrden.Recordset.RecordCount = 0 Then tCreditos = 0: GoTo Cuota

DtOrden.Recordset.MoveFirst
   
   Do While Not DtOrden.Recordset.EOF
        Tord = Tord + DtOrden.Recordset("Saldo")
        tCreditos = Tord
        
        DtOrden.Recordset.MoveNext
     
   Loop
   
Cuota:
'Comienza a sumar Cuota  Social
If DtCtaSoc.Recordset.RecordCount = 0 Then tCuota = 0: GoTo AyudaSoc

DtCtaSoc.Recordset.MoveFirst
    
    Do While Not DtCtaSoc.Recordset.EOF
        tcuot = tcuot + DtCtaSoc.Recordset("Saldo")
        tCuota = tcuot
        
         DtCtaSoc.Recordset.MoveNext

    Loop
    

AyudaSoc:
'Comienza sumar  Ayuda Social
If DtAyuda.Recordset.RecordCount = 0 Then totAyuda = 0: GoTo Suma

DtAyuda.Recordset.MoveFirst

    Do While Not DtAyuda.Recordset.EOF
        Tayuda = Tayuda + DtAyuda.Recordset("Saldo")
        totAyuda = Tayuda
        
         DtAyuda.Recordset.MoveNext
    Loop
    
Suma:
s(0) = tCreditos
s(1) = tCuota
s(2) = totAyuda


Total = s(0) + s(1) + s(2)


End Sub

 



Private Sub btnConfirma_Click()
Dim msg As String
  msg = MsgBox("Confirma el pago de la Orden N°  " & DtOrden.Recordset("Ord_Nro"), vbYesNo + vbQuestion)
  
  If msg = vbYes Then
    DtOrden.Recordset.Edit
        DtOrden.Recordset("Saldo") = 0
    DtOrden.Recordset.Update
    
     'Pasa los datos a la tabla pagos
    With DtPagos
    .Recordset.AdelantoNew
    .Recordset("Fecha") = Fecha
    .Recordset("Cliente") = DtOrden.Recordset("IDCliente")
    .Recordset("Comercio") = DtOrden.Recordset("comercio")
    .Recordset("Valor") = DtOrden.Recordset("Valor")
    .Recordset("Saldo") = DtOrden.Recordset("Saldo")
    .Recordset("Mes") = DtOrden.Recordset("Mes")
    .Recordset("Ano") = DtOrden.Recordset("Ano")
    .Recordset("Cta") = DtOrden.Recordset("Cuota")
    .Recordset("NroOrd") = DtOrden.Recordset("Ord_Nro")
    .Recordset("Int") = DtOrden.Recordset("Int")
    .Recordset("Situacion") = DtOrden.Recordset("Situacion")
    .Recordset("Rec") = 0
    .Recordset("RegNo") = 0
    .Recordset.Update
    End With
    
       MsgBox "Pago Registrado con Exito", vbInformation
       
    DtOrden.Refresh
    DBGrid1.Refresh
    tot
       
         Else
      'Cancel = -1
      Exit Sub
  End If
         
    
            
    
End Sub


Private Sub btnHaceEntrega_Click()
 Ent = InputBox("Ingrese el valor de la Entrega", "Cuadro para el ingreso de Entregas")
 
 Dim v As Single, sd As Single
        
EntCta
    
End Sub

Private Sub btnPagaTodo_Click()
Dim msg As String
  msg = MsgBox("Confirma el pago de Todas las Ordenes + Ayuda Soc y Cta? ", vbYesNo + vbQuestion)
  
  If msg = vbYes Then
   
    
     'Pasa los datos a la tabla pagos
With DtPagos
If DtOrden.Recordset.RecordCount = 0 Then GoTo cta
  DtOrden.Recordset.MoveFirst 'posiciona en el primer registro

Do While Not DtOrden.Recordset.EOF
   DtOrden.Recordset.Edit
        DtOrden.Recordset("Saldo") = 0
    DtOrden.Recordset.Update
  
  
    
    .Recordset.AdelantoNew
    .Recordset("Fecha") = Fecha
    .Recordset("Cliente") = DtOrden.Recordset("IDCliente")
    .Recordset("Comercio") = DtOrden.Recordset("comercio")
    .Recordset("Valor") = DtOrden.Recordset("Valor")
    .Recordset("Saldo") = DtOrden.Recordset("Saldo")
    .Recordset("Mes") = DtOrden.Recordset("Mes")
    .Recordset("Ano") = DtOrden.Recordset("Ano")
    .Recordset("Cta") = DtOrden.Recordset("Cuota")
    .Recordset("NroOrd") = DtOrden.Recordset("Ord_Nro")
    .Recordset("Int") = DtOrden.Recordset("Int")
    .Recordset("Situacion") = DtOrden.Recordset("Situacion")
    .Recordset("Rec") = 0
    .Recordset("RegNo") = 0
    .Recordset.Update
               DtOrden.Recordset.MoveNext
Loop
    End With
    
MsgBox "Todas la ordenes fueron Registradas", vbInformation
    
       
    DtOrden.Refresh

tot
cta:

If DtCtaSoc.Recordset.RecordCount = 0 Then GoTo ayuda
  DtCtaSoc.Recordset.MoveFirst 'posiciona en el primer registro

'COMIENZO DE  TRABAJO CON LA TABLA CTASOC

Do While Not DtCtaSoc.Recordset.EOF
   
  

   
   DtCtaSoc.Recordset.Edit
        DtCtaSoc.Recordset("Saldo") = 0
        DtCtaSoc.Recordset("Pago") = "S"
    DtCtaSoc.Recordset.Update
    
    DtCtaSoc.Recordset.MoveNext

Loop

tot
ayuda:

      If DtAyuda.Recordset.RecordCount = 0 Then Exit Sub
  DtAyuda.Recordset.MoveFirst 'posiciona en el primer registro


Do While Not DtAyuda.Recordset.EOF
   
   
   DtAyuda.Recordset.Edit
        DtAyuda.Recordset("Saldo") = 0
    DtAyuda.Recordset.Update
   
   DtAyuda.Recordset.MoveNext
Loop
   tot
         
         Else
      'Cancel = -1
      Exit Sub
  End If

End Sub



Private Sub Command1_Click()
AdelantoCta
End Sub


Private Sub DBGrid1_DblClick()
    MsgBox DtOrden.Recordset("Valor") & DtOrden.Recordset("Mes")

End Sub

Private Sub Fecha_Change()
    Fecha.SelStart = 0
    Fecha.SelLength = Len(Fecha.Text)
End Sub

Sub AdelantoCta()
Dim Adelanto As Double
Adelanto = InputBox("Ingrese el adelanto de la cuota")
If Adelanto = "" Then Exit Sub

If DtCtaSoc.Recordset.RecordCount = 0 Then Adelanta: Exit Sub
  DtCtaSoc.Recordset.MoveFirst 'posiciona en el primer registro

'COMIENZO DE  TRABAJO CON LA TABLA CTASOC

Do While Not DtCtaSoc.Recordset.EOF
   
  
'si el valor de la entrega es sup al saldo  saldo = 0 (paga cuota)
  
If Adelanto >= DtCtaSoc.Recordset("saldo") Then
 
Adelanto = Adelanto - DtCtaSoc.Recordset("Saldo")

MsgBox "Valor de Adelanto despues de restar saldo $10  " & Adelanto
   DtCtaSoc.Recordset.Edit
        DtCtaSoc.Recordset("Saldo") = 0
        DtCtaSoc.Recordset("Pago") = "S"
    DtCtaSoc.Recordset.Update
  

  
        
 End If
'>>>>>>>>>>>>>>>Si lo que queda de la entrega no cubre el saldo<<<<<
If Adelanto < DtCtaSoc.Recordset("Saldo") Then
MsgBox "El adelanto no supera esta cuota"
     sd = DtCtaSoc.Recordset("Saldo") - Adelanto
      
      DtCtaSoc.Recordset.Edit
       
        
        DtCtaSoc.Recordset("Saldo") = sd
      
      DtCtaSoc.Recordset.Update
    
    DtCtaSoc.Refresh
    
    tot
               
               
               Exit Do
End If
 
    DtCtaSoc.Recordset.MoveNext

 Loop
 MsgBox "EOF DE CTA Adelanto= " & Adelanto
Adelanta
End Sub
Private Sub Form_Load()
    Fecha = Format(Date, "dd/mm/yy")
    txtMes = Month(Date)
    txtAño = Year(Date)
    
    tCuota.Text = 0
    tCreditos.Text = 0
    Total.Text = 0
    totAyuda = 0
    ValCta = 45
    
End Sub

Private Sub txtAño_GotFocus()
    txtAño.SelStart = 0
    txtAño.SelLength = Len(txtAño.Text)
End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then



With DtOrden
            .RecordSource = "Select IDCliente,  Comercio, Nomb_Com, Valor, Saldo, Mes, Ano, Cuota,Tot_Cuot, Ord_Nro,Int, Situacion, Rec, " _
            & "FEmision FROM TBL_Ordenes " _
            & "WHERE IDCliente = " & txtCodSoc.Text & "" _
            & "And Mes<= " & txtMes & "" _
            & "And Ano = " & txtAño & "" _
            & "And Saldo <> 0 "
        .Refresh
        
        End With
        
        tot
End If
End Sub

Private Sub txtCodSoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 
    DtSocio.Recordset.Index = "IndNroSoc"
    DtSocio.Recordset.Seek "=", txtCodSoc
    
      If DtSocio.Recordset.NoMatch Then
            MsgBox "Cliente Inexistente", vbQuestion
            Exit Sub
         Else
         txtNombre = DtSocio.Recordset("Apellido") & "  " & DtSocio.Recordset("Nombre")
      End If
         
    
Fecha.Enabled = True
txtAño.Enabled = True
txtMes.Enabled = True
txtMes.SetFocus

End If
End Sub

Private Sub txtMes_GotFocus()
    txtMes.SelStart = 0
    txtMes.SelLength = Len(txtMes.Text)
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    tCuota.Text = 0
    tCreditos.Text = 0
    totAyuda.Text = 0
    Total.Text = 0

With DtOrden
            .RecordSource = "Select IDCliente,  Comercio, Nomb_Com, Valor, Saldo, Mes, Ano, Cuota,Tot_Cuot, Ord_Nro,Int, Situacion, Rec, " _
            & "FEmision FROM TBL_Ordenes " _
            & "WHERE IDCliente = " & txtCodSoc.Text & "" _
            & "And Mes<= " & txtMes & "" _
            & "And Ano = " & txtAño & "" _
            & "And Saldo <> 0 "
        .Refresh
End With

With DtCtaSoc
    .RecordSource = "Select * FROM CtaSoc " _
            & "WHERE IDCliente = " & txtCodSoc.Text & "" _
            & "And Mes<= " & txtMes & "" _
            & "And Ano = " & txtAño & "" _
            & "And Saldo <> 0 "
            .Refresh
End With

With DtAyuda
            .RecordSource = "Select IDCliente, Valor, Saldo, Mes, " _
            & "Ano FROM Ayuda " _
            & "WHERE IDCliente = " & txtCodSoc.Text & "" _
            & "And Mes<= " & txtMes & "" _
            & "And Ano = " & txtAño & "" _
            & "And Saldo <> 0 "
        .Refresh
End With

tot
End If

 
End Sub

Sub FiltraOrden()
With DtOrden
            .RecordSource = "Select IDCliente,  Comercio, Nomb_Com, Valor, Saldo, Mes, Ano, Cuota,Tot_Cuot, Ord_Nro, Situacion, Rec, " _
            & "FEmision FROM TBL_Ordenes " _
            & "WHERE IDCliente = " & txtCodSoc.Text & "" _
            & "And Mes<= " & txtMes & "" _
            & "And Ano = " & txtAño & "" _
            & "And Saldo <> 0 "
        .Refresh
End With
End Sub
