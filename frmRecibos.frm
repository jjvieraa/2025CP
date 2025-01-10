VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjRecibo 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir RECIBOS"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "Ver Tabla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """##M####"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Generar"
      DownPicture     =   "frmRecibos.frx":0442
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2220
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1508
      _Version        =   393216
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
            LCID            =   14346
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
            LCID            =   14346
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
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Sit.Laboral:"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Vto.Presup:"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Cobrador:"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "fjRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim WithEvents rstT As ADODB.Recordset
Attribute rstT.VB_VarHelpID = -1
Dim WithEvents rst2 As ADODB.Recordset
Attribute rst2.VB_VarHelpID = -1


Const kNroCobrador = 0
Const kFechaVtoPresup = 1
    
Dim cREC As clsRecibos


'=====================================================================
Private Sub Form_Load()
'=====================================================================
    
    
   'Set adoconn = New ADODB.Connection
   'adoconn.CursorLocation = adUseClient
   'adoconn.Open "dsn=JIMMY"
    Set cREC = New clsRecibos
    
    cREC.m0fTomaParametros

    Text1(kFechaVtoPresup).MaxLength = 10
    Text1(kNroCobrador).MaxLength = 3
    Text1(kFechaVtoPresup).Text = CStr(vpnPrspHst) & "/" & _
        Right(vptMesPresup, 2) & "/" & Left(vptMesPresup, 4)
   'Text1(kNroCobrador).SetFocus
    msCargaComboSituacLaboral Combo1
    
End Sub


'======================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
'======================================================
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"        ' COMO SI PULSARA ENTER
    End If
End Sub


'=====================================================================
Private Sub Text1_change(Index As Integer)
'=====================================================================
Select Case Index
    Case kFechaVtoPresup        'Formato 10/mm/aaaa
        If Len(Text1(kFechaVtoPresup).Text) = 2 Then
            Text1(kFechaVtoPresup).Text = Text1(kFechaVtoPresup).Text & "/"
            Text1(kFechaVtoPresup).SelStart = 3
        End If
End Select
End Sub







'=====================================================================
Private Sub Text1_GotFocus(Index As Integer)
'=====================================================================
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub



Private Sub Text1_LostFocus(Index As Integer)
'pasa al siguiente campo: Fecha Vto
If Index = kNroCobrador Then
    Text1(kFechaVtoPresup).SetFocus
End If
End Sub

'=====================================================================
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
'=====================================================================

Select Case Index
    Case kNroCobrador  'recibo
        If Not IsNumeric(Text1(kNroCobrador).Text) Then
            Cancel = True
        End If
    Case kFechaVtoPresup
       If Not IsDate(Text1(kFechaVtoPresup).Text) Then
            Cancel = True
       Else
            Dim yMome As Integer
            yMome = Year(CDate(Text1(kFechaVtoPresup).Text))
            If yMome < kMenorAñoTrabajo Or _
                yMome > kMayorAñoTrabajo Then
                Cancel = True
            End If
       End If
       
End Select
End Sub


'=====================================================================
Private Sub Command1_Click()
'=====================================================================
    Dim nM As Integer
    
    If Check1.Value = vbChecked Then
      DataGrid1.Visible = True
    End If
    If mfEstaVacio(Text1(kNroCobrador).Text) Then
        Text1(kNroCobrador).SetFocus
        Exit Sub
    End If
    If mfEstaVacio(Text1(kFechaVtoPresup).Text) Then
        Text1(kFechaVtoPresup).SetFocus
        Exit Sub
    End If
    mAviso ("Espere por favor...")
    If cREC.m1fCargaLosRegistrosAImprimir(CDate(Text1(kFechaVtoPresup).Text), _
        CInt(Text1(kNroCobrador).Text), mfDevCodSitLaboral(Combo1.Text)) Then
        cREC.m2fImprimeRecibos
    End If
    'Set DataGrid1.DataSource = rst2

End Sub



