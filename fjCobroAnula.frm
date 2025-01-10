VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjCobroAnula 
   Caption         =   "Anula Un Cobro"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   570
      Width           =   5055
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3201
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
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Motivo por el que se anula:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "No Recibo:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "fjCobroAnula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoA As New ADODB.Recordset
Dim adoPagos As New ADODB.Recordset
Dim adoPagos2 As New ADODB.Recordset
Dim adoOrdenes As New ADODB.Recordset
Dim cOrd As New clsOrdenes


Private Sub cmdAnular_Click()
Label2.Caption = "Agregando registros en Pagos"
Label2.Refresh
If adoPagos2.State = adStateOpen Then adoPagos2.Close
adoPagos2.Open "SELECT * FROM tbl_Pagos;", adoconn, adOpenDynamic, adLockOptimistic

Dim nk As Integer
adoPagos.MoveFirst
Do While Not adoPagos.EOF
    adoPagos2.AddNew
    For nk = 0 To adoPagos.Fields.Count - 2 'el ultimo no, por que es autonumerico
        adoPagos2(nk) = adoPagos(nk)
    Next nk
    adoPagos2("pag_valor") = adoPagos2("pag_valor") * (-1)
    adoPagos2("pag_motivo") = 8
    adoPagos2.Update
    adoPagos.MoveNext
Loop


Label2.Caption = "Agregando pagos en Ordenes"
Label2.Refresh
'If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
'adoOrdenes.Open "SELECT * FROM tbl_Ordenes order BY ord_NroOrden;", adoconn, adOpenDynamic, adLockOptimistic

Dim lNroOrden As Long
Dim dVto As Date
Dim lNroSoc As Long

cOrd.msInicia2
adoPagos.MoveFirst
lNroSoc = adoPagos("pag_NroSoc")
Do While Not adoPagos.EOF
    lNroOrden = adoPagos("pag_NroOrden")
   If lNroOrden = 1 Or lNroOrden = 2 Then
        'ubica la orden
        dVto = CDate(Right(adoPagos("pag_det"), 10))
        If Not cOrd.fBuscaUnaCuotaOAyuda(lNroSoc, lNroOrden, dVto) Then
            MsgBox "Error 745: "
        End If
    
    Else    'es una orden comun
        'ubica la orden
        If Not cOrd.fBuscaUnaOrden(lNroOrden) Then
        MsgBox "Error 746:"
        End If
    End If
    Set adoOrdenes = cOrd.adoOrdenes
    If Not adoOrdenes.EOF Then
        'SI ESTA CERRADA LA ABRE
        If Year(adoOrdenes("ord_cerro")) > 1900 Then
            adoOrdenes("ord_cerro") = CDate("01/01/1900")
        End If
        'SI ES CUOTA
        If adoPagos("pag_valor") = adoOrdenes("ord_cuota") And _
            adoPagos("pag_mon") = "P" And adoOrdenes("ord_ctasPagas") > 0 Then
                adoOrdenes("ord_ctasPagas") = adoOrdenes("ord_ctasPagas") - 1
        ElseIf adoPagos("pag_valor") = adoOrdenes("ord_MEcuota") And _
            Not adoPagos("pag_mon") = "P" And adoOrdenes("ord_ctasPagas") > 0 Then
                adoOrdenes("ord_ctasPagas") = adoOrdenes("ord_ctasPagas") - 1
        'es ent cta
        Else
                adoOrdenes("ord_entcta") = adoOrdenes("ord_entcta") - adoPagos("pag_valor")
        End If
        adoOrdenes.Update
    Else
        MsgBox "ERROR 645"
    End If
    adoPagos.MoveNext
Loop
adoPagos.MoveFirst
Label2.Caption = "Agregando registro en Accion"
Label2.Refresh
msRegistraUnaAccion 8, _
                    adoPagos("pag_NroPago"), _
                    "Anula Cobro: " & Text2.Text, _
                    vpnFuncionario, _
                    Format(Date, "short date"), _
                    Format(Time, "short time")

Unload Me
End Sub

Private Sub cmdVer_Click()
    If Len(Trim(Text2.Text)) = 0 Then Exit Sub
    Set adoA.ActiveConnection = adoconn
    
    If adoA.State = adStateOpen Then adoA.Close
    
    adoA.Open "SELECT * FROM tbl_Pagos WHERE pag_NroPago =" & Text1.Text & ";", adoconn, adOpenDynamic, adLockOptimistic
    If adoA.RecordCount > 0 Then
        Set adoPagos = adoA
        Set DataGrid1.DataSource = adoPagos
        cmdAnular.Enabled = True
    End If
End Sub

Private Sub Form_Load()
cmdAnular.Enabled = False
Label2.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If adoPagos.State = adStateOpen Then adoPagos.Close
    Set adoPagos = Nothing
    If adoPagos2.State = adStateOpen Then adoPagos2.Close
    Set adoPagos2 = Nothing
    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    Set adoOrdenes = Nothing
    If adoA.State = adStateOpen Then adoA.Close
    Set adoA = Nothing
    Set cOrd = Nothing
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Not IsNumeric(Text1.Text) Then Cancel = True
End Sub

