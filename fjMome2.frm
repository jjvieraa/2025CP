VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fjMome2 
   Caption         =   "Actualizando fecha..."
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "revisa "
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Acción: Actualiza la fecha de emisión de los conformes importados de la base anterior con fecha 1/1/1900"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "fjMome2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim adoA As New ADODB.Recordset
Dim adoPagos As New ADODB.Recordset
Dim adoPagos2 As New ADODB.Recordset
Dim adoOrdenes As New ADODB.Recordset


Private Sub otro()
    Dim sM As String
    Screen.MousePointer = vbHourglass
    Set adoPagos.ActiveConnection = adoconn
    Label2.Caption = "Abriendo...."
    Label2.Refresh
    If adoA.State = adStateOpen Then adoA.Close
     sM = "select * FROM tbl_Pagos INNER JOIN tbl_Socios " & _
            "ON tbl_Socios.NroSoc = tbl_Pagos.pag_NroSoc " & _
            "WHERE (pag_Motivo = 5 OR pag_Motivo = 6 OR pag_Motivo = 7) " & _
            "AND (pag_Fecha BETWEEN #1/1/3# AND #1/31/3#) " & _
            "AND (CodSitLab = 3 OR CodSitLab = 4);"

    adoA.Open sM, adoconn, adOpenDynamic, adLockOptimistic
    If adoA.RecordCount < 1 Then
        MsgBox "No encuentra registros...."
        Unload Me
        Exit Sub
    End If

   Set adoPagos = adoA
 
   If adoPagos2.State = adStateOpen Then adoPagos2.Close
     
    adoPagos2.Open "SELECT * FROM TBL_PAGOS;", adoconn, adOpenDynamic, adLockOptimistic
      pb.Max = adoPagos.RecordCount
    Dim tP As Long
    adoPagos.MoveFirst
    tP = adoPagos.RecordCount
     Do While Not adoPagos.EOF
        pb.Value = adoPagos.AbsolutePosition
        Label2.Caption = "Eliminando...." & tP
        Label2.Refresh
        tP = tP - 1
        adoPagos2.MoveFirst
        adoPagos2.Find ("pag_auto =" & adoPagos!pag_auto)
        If Not adoPagos2.EOF Then
            adoPagos2.Delete
            adoPagos2.Update
        Else
            MsgBox "No pudo eliminar en tbl_pagos el reg " & adoPagos!pag_auto
        End If
        adoPagos.MoveNext
     Loop

   

    Dim lNroOrden As Long
    Dim dVto As Date
    Dim lNroSoc As Long
    Dim mBien As Boolean
    
    pb.Max = adoPagos.RecordCount
    adoPagos.MoveFirst
    tP = adoPagos.RecordCount
    Do While Not adoPagos.EOF
        
        Label2.Caption = "Actualizando...." & tP
        Label2.Refresh
        tP = tP - 1
        pb.Value = adoPagos.AbsolutePosition
        
        lNroSoc = adoPagos("pag_NroSoc")
        lNroOrden = adoPagos("pag_NroOrden")
       'If lNroSoc = 103 And lNroOrden = 1 Then
       '     Debug.Print lNroSoc
       'End If
       If lNroOrden = 1 Or lNroOrden = 2 Then
            'ubica la orden
            dVto = CDate(Right(adoPagos("pag_det"), 10))
            If Not fBuscaUnaCuotaOAyuda(lNroSoc, lNroOrden, dVto) Then
                MsgBox "Error 745: No encuentra Orden No:" & lNroOrden & " del Cliente: " & lNroSoc
                mBien = False
            Else
                mBien = True
            End If
        
        Else    'es una orden comun
            'ubica la orden
            If Not fBuscaUnaOrden(lNroOrden) Then
                MsgBox "Error 746: No encuentra orden No: " & lNroOrden & " del Cliente: " & lNroSoc
                mBien = False
            Else
                mBien = True
            End If
        End If
        
        
        If mBien And Not adoOrdenes.EOF Then
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
         End If
        adoPagos.MoveNext
    Loop
    Unload Me
End Sub

Private Sub cmdVer_Click()
End Sub


Private Sub Command1_Click()
   
    Dim sMom As String
    Dim nMom As Long
    Screen.MousePointer = vbHourglass
    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT * FROM tbl_ordenes;"
    adoOrdenes.Open sMom, adoconn, adOpenDynamic, adLockOptimistic
    adoOrdenes.MoveFirst
        pb.Visible = True
        pb.Min = 0
        pb.Max = adoOrdenes.RecordCount
    Do While Not adoOrdenes.EOF
        If adoOrdenes.AbsolutePosition Mod 100 = 0 Then
            Label3.Caption = adoOrdenes.AbsolutePosition & "  " & nMom
            Label3.Refresh
                pb.Value = adoOrdenes.AbsolutePosition
        End If
        If adoOrdenes!ord_FEmis < #1/1/1930# Then
            Debug.Print adoOrdenes!ord_FEmis
            adoOrdenes!ord_FEmis = adoOrdenes!ord_FVto - 30
            adoOrdenes.Update
            Debug.Print adoOrdenes!ord_FEmis
            nMom = nMom + 1
        End If
        adoOrdenes.MoveNext
    Loop
    Screen.MousePointer = vbDefault
    adoOrdenes.Close
    Set adoOrdenes = Nothing
    Unload Me
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
End Sub



'=====================================================================
Private Function fBuscaUnaCuotaOAyuda(lSoc As Long, lOrd As Long, _
fVto As Date) As Boolean
'=====================================================================
    
    Dim sMom As String
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT * FROM tbl_ordenes " & _
        "WHERE ORD_NroSoc =" & CStr(lSoc) & _
        " AND ord_NroOrden =" & CStr(lOrd) & _
        " AND ord_fVto =#" & mfInvierteMes(CStr(fVto)) & "#;"
    adoOrdenes.Open sMom, adoconn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount = 0 Then
        fBuscaUnaCuotaOAyuda = False
        Exit Function
     ElseIf Not adoOrdenes.RecordCount = 1 Then
        fBuscaUnaCuotaOAyuda = False
        Exit Function
    End If
    
   fBuscaUnaCuotaOAyuda = True
End Function














'================================================================
Private Sub Command2_Click()
        Screen.MousePointer = vbHourglass
        
        
        If adoA.State = adStateOpen Then adoA.Close
        adoA.Open "SELECT * FROM tbl_prepago;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
        
        adoA.MoveFirst
        pb.Visible = True
        pb.Min = 0
        pb.Max = adoA.RecordCount
        Do While Not adoA.EOF
                pb.Value = adoA.AbsolutePosition
                Label3.Caption = adoA.AbsolutePosition
                Label3.Refresh
                If adoA!pp_Presup = "20031" Then
                    adoA!pp_Presup = "200301"
                    adoA.Update
                End If
                adoA.MoveNext
        Loop
        If adoA.State = adStateOpen Then adoA.Close
        
        Set adoA = Nothing
        Screen.MousePointer = vbDefault
        
        Unload Me
End Sub

'=====================================================================
Private Function fBuscaUnaOrden(lPrm As Long) As Boolean
'=====================================================================
    
    Dim sMom As String
    

    If adoOrdenes.State = adStateOpen Then adoOrdenes.Close
    sMom = "SELECT * FROM tbl_ordenes " & _
        "WHERE ORD_NroOrden =" & CStr(lPrm) & ";"
    adoOrdenes.Open sMom, adoconn, adOpenDynamic, adLockOptimistic
    If adoOrdenes.RecordCount <> 1 Then
        fBuscaUnaOrden = False
        Exit Function
    End If
    
   fBuscaUnaOrden = True
End Function

Private Sub Command5_Click()
    Screen.MousePointer = vbHourglass

    If adoA.State = adStateOpen Then adoA.Close
    adoA.Open "SELECT  * FROM tbl_socios ORDER BY nrosoc;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

     adoA.MoveFirst
    Do While Not adoA.EOF
        Label3.Caption = adoA!NroSoc & "   " & CInt(adoA!NroCob)
        Label3.Refresh
        adoA.MoveNext
    Loop

    Screen.MousePointer = vbDefault

    If adoA.State = adStateOpen Then adoA.Close

    Set adoA = Nothing


End Sub
