VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMisInformes 
   Caption         =   "Informes II"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Resume la tbl_Pagos  en tbl_JI2"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Suma el total de cuotas ."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resume la tbl_Ordenes en tbl_JI1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   6495
   End
   Begin VB.Label lblReg 
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "frmMisInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoA As New ADODB.Recordset
Dim adoB As New ADODB.Recordset
Dim adoC As New ADODB.Recordset


Dim vAM As String           'año mes  aaaamm



'===============================
Private Sub Command1_Click()
'===============================
Label1.Caption = "Abriendo tbl_Ordenes...."
Label1.Refresh
adoA.Open "select * FROM tbl_Ordenes ORDER BY ORD_FEmis;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText


Label1.Caption = "Creando tabla virtual..."
Label1.Refresh
'crea una tabla virtual
Set adoB.ActiveConnection = adoconn
Set adoB = New ADODB.Recordset
adoB.Fields.Append "fAM", adChar, 6
adoB.Fields.Append "fMes", adChar, 2
adoB.Fields.Append "fAno", adChar, 4

adoB.Fields.Append "fCntsCuoP", adInteger, 2  'Cuotas Pesos
adoB.Fields.Append "fValCuoP", adDouble       'Cuotas Pesos
adoB.Fields.Append "fCntsCuoD", adInteger, 2  'Cuotas Dolar
adoB.Fields.Append "fValCuoD", adDouble       'Cuotas Dolar OJO : TODOS LOS VALORES EN PESOS
adoB.Fields.Append "fCntsCuoU", adInteger, 2  'Cuotas UR
adoB.Fields.Append "fValCuoU", adDouble       'Cuotas UR

adoB.Fields.Append "fCntsAyuP", adInteger, 2  'Ayuda Pesos
adoB.Fields.Append "fValAyuP", adDouble       'Ayuda Pesos
adoB.Fields.Append "fCntsAyuD", adInteger, 2  'Ayuda Dolar
adoB.Fields.Append "fValAyuD", adDouble       'Ayuda Dolar
adoB.Fields.Append "fCntsAyuU", adInteger, 2  'Ayuda UR
adoB.Fields.Append "fValAyuU", adDouble       'Ayuda UR

adoB.Fields.Append "fCntsCarP", adInteger, 2  'carniceria Pesos
adoB.Fields.Append "fValCarP", adDouble       'carniceria Pesos
adoB.Fields.Append "fCntsCarD", adInteger, 2  'carniceria Dolar
adoB.Fields.Append "fValCarD", adDouble       'carniceria Dolar
adoB.Fields.Append "fCntsCarU", adInteger, 2  'carniceria UR
adoB.Fields.Append "fValCarU", adDouble       'carniceria UR

adoB.Fields.Append "fCntsvalP", adInteger, 2  'vales Pesos
adoB.Fields.Append "fValvalP", adDouble       'vales Pesos
adoB.Fields.Append "fCntsvalD", adInteger, 2  'vales Dolar
adoB.Fields.Append "fValvalD", adDouble       'vales Dolar
adoB.Fields.Append "fCntsvalU", adInteger, 2  'vales UR
adoB.Fields.Append "fValvalU", adDouble       'vales UR

adoB.Fields.Append "fCntsordP", adInteger, 2  'ordenes Pesos
adoB.Fields.Append "fValordP", adDouble       'ordenes Pesos
adoB.Fields.Append "fCntsordD", adInteger, 2  'ordenes Dolar
adoB.Fields.Append "fValordD", adDouble       'ordenes Dolar
adoB.Fields.Append "fCntsordU", adInteger, 2  'ordenes UR
adoB.Fields.Append "fValordU", adDouble       'ordenes UR


adoB.Fields.Append "fCntVacio", adInteger, 2  'sin uso
adoB.Fields.Append "fValMe", adDouble    'Valor en P de las ME

adoB.CursorType = adOpenDynamic
adoB.LockType = adLockOptimistic
adoB.Open

'------------------------------------
'borra los registros de la tabla tbl_JI1
'------------------------------------
Label1.Caption = "Limpiando..."
Label1.Refresh
pb.Visible = True
pb.Min = 0
pb.Max = adoA.RecordCount
adoC.Open "delete * FROM tbl_JI1;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
adoC.Open "select * FROM tbl_JI1;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText





'------------------------------------
'recorre la tabla tbl_Ordenes y
'guarda un resumen en la tabla tbl_JI1
'------------------------------------
Label1.Caption = "Recorriendo..."
Label1.Refresh
Dim nM, nf As Long
nf = adoA.RecordCount
pb.Min = 0
pb.Max = nf

'ordena la tabla
adoB.Sort = "fam"

'recorre tbl_ordenes y guarda en tbl_ji1

adoA.MoveFirst
Do While Not adoA.EOF
    If nM Mod 100 Then
        pb.Value = nM
        lblReg.Caption = nf - nM
        lblReg.Refresh
    End If
    'si NO es una orden anulada
    If Not adoA!ord_tipo = 4 Then
        msTomaVAM (adoA!ord_FEmis)
        If mfNoEncuentraVam Then
            msAgregaReg
        End If
        msGuardaDatosEnAdo
        'Debug.Print adoB!fam, adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU
    End If
    adoA.MoveNext
    nM = nM + 1
Loop


   'Set fjMome.DataGrid1.DataSource = adoB
   ' fjMome.Show
   ' MsgBox "Probando"




'------------------------------------
'guarda en tabla
'------------------------------------
Label1.Caption = "Guardando...."
Label1.Refresh
pb.Min = 0
pb.Max = adoB.RecordCount
pb.Value = 0
adoB.MoveFirst
Do While Not adoB.EOF
   If adoB.AbsolutePosition Mod 100 Then pb.Value = adoB.AbsolutePosition
    'Debug.Print adoB!fam, adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU
    msGuardaReg
    adoB.MoveNext
Loop
adoC.Update

'------------------------------------
'formatea los valores decimales de la tabla tbl_JI1
'------------------------------------
Label1.Caption = "Formateando...."
Label1.Refresh
pb.Min = 0
pb.Max = adoC.RecordCount
pb.Value = 0
adoC.MoveFirst
Do While Not adoC.EOF
    pb.Value = adoC.AbsolutePosition
    'Debug.Print adoB!fam, adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU
    msFormateaReg
    adoC.MoveNext
Loop




'------------------------------------
'libera
'------------------------------------
Label1.Caption = ""
Label1.Refresh

adoA.Close
Set adoA = Nothing
adoB.Close
Set adoB = Nothing
adoC.Close
Set adoC = Nothing
pb.Visible = False

End Sub


'===============================
Private Sub msFormateaReg()
'===============================
    'deja con 2 decimales unicamente
    
    Dim mS As String
    Dim nc As Integer
    'solo los valores doubles
    For nc = 4 To 34 Step 2
        mS = adoC(nc)
        mS = mfCambiaPuntoPorComa(mS)
        mS = Format(mS, "0.00")
        adoC(nc) = CDbl(mS)
    Next
    adoC.Update

End Sub


'===============================
Private Sub msGuardaReg()
'===============================
    Dim ni As Integer
    adoC.AddNew
    For ni = 0 To 34
        adoC(ni) = adoB(ni)
        'Debug.Print ni, adoB(ni)
    Next
    'adoC!fCntsordP = adoB!fCntsordP
    'adoC!fValordP = adoB!fValordP
    adoC.Update
    'Debug.Print adoB!fam, adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU
End Sub


'===============================
Private Sub msGuardaDatosEnAdo()
'==============================
    'CUOTA
    If adoA!ord_NroOrden = 1 Then
        msGuardaDatos2 adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fValME
    
    'AYUDA
    ElseIf adoA!ord_NroOrden = 2 Then
        msGuardaDatos2 adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU, adoB!fValME
    
    Else
        'CARNICERIA
        If adoA!ORD_NroCom = 115 Then
            msGuardaDatos2 adoB!fCntscarP, adoB!fValcarP, adoB!fCntscarD, adoB!fValcarD, adoB!fCntscarU, adoB!fValcarU, adoB!fValME
    
        'VALES
        ElseIf adoA!ORD_NroCom = 190 Then
            msGuardaDatos2 adoB!fCntsvalP, adoB!fValvalP, adoB!fCntsvalD, adoB!fValvalD, adoB!fCntsvalU, adoB!fValvalU, adoB!fValME
    
        'ORDENES
        Else
            msGuardaDatos2 adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fValME
        End If
    End If
   'Debug.Print adoB(0), adoB.Fields(4).Value
   'Set fjMome.DataGrid1.DataSource = adoB
   'fjMome.Show
   adoB.Update
End Sub


'===============================
Private Sub msGuardaDatos2(CantP As Field, ValP As Field, CantD As Field, ValD As Field, CantU As Field, ValU As Field, ValME As Field)
'===============================
    'es en pesos
    If adoA!ord_Mon = "P" Then
        CantP = CantP + 1
        ValP = ValP + adoA!ord_cuota * adoA!ORD_PLAN
    
    'es en dolares
    ElseIf adoA!ord_Mon = "D" Then
        CantD = CantD + 1
        ValD = ValD + adoA!ord_cuota * adoA!ORD_PLAN
        ValME = ValME + adoA!ord_mecuota * adoA!ORD_PLAN
    'es en UR
    Else
        CantU = CantU + 1
        ValU = ValU + adoA!ord_cuota * adoA!ORD_PLAN
        ValME = ValME + adoA!ord_mecuota * adoA!ORD_PLAN
    End If
    'Debug.Print CantP, ValP
End Sub




'===============================
Private Sub msAgregaReg()
'===============================
    adoB.AddNew
    adoB!fam = vAM
    adoB!fMes = Right(vAM, 2)
    adoB!fAno = Left(vAM, 4)
    adoB.Update
End Sub



'===============================
Private Function mfNoEncuentraVam() As Boolean
'===============================
    adoB.Find ("fam ='" & vAM & "'")
    If adoB.EOF Then
        mfNoEncuentraVam = True
    Else
        mfNoEncuentraVam = False
    End If
End Function


'===============================
Private Sub msTomaVAM(nPrm As Date)
'===============================
'Toma la fecha de emisión
'si dia > 10 entonces mes++, si mes=13 entonces mes=1;año++
    
    Dim mDia, mMes As Byte
    Dim mAnio As Integer
    
    mDia = Day(nPrm)
    mMes = Month(nPrm)
    mAnio = Year(nPrm)
    
    If mDia > 10 Then
        mMes = mMes + 1
        If mMes = 13 Then
            mMes = 1
            mAnio = mAnio + 1
        End If
    End If
    
    If mMes < 10 Then
            vAM = mAnio & "0" & mMes
    Else
            vAM = mAnio & mMes
    End If

End Sub


'adoOrdenes!ord_nrosoc = vlNroSoc
'adoOrdenes!ORD_NroCom = vlNroComerc
'adoOrdenes!ord_NroOrden = vlNroOrden
'adoOrdenes!ORD_DEPEND = vlNroDepend
'doOrdenes!ord_cuota = vnsCuota
'adoOrdenes!ord_Femis = vdFEmis
'adoOrdenes!ord_FVto = vdFVto
'adoOrdenes!ORD_PLAN = vnPlan
'adoOrdenes!ord_ctasPagas = vnCtasPaga
'adoOrdenes!ord_EntCta = vnsEntCta
'adoOrdenes!ord_Recarg = vnsRecarg
'adoOrdenes!ord_Mon = vsMoneda
'adoOrdenes!ord_mecuota = vnsMECuota
'adoOrdenes!ORD_MEPagos = vnsMEPagos
'adoOrdenes!ord_cerro = vdCerro                 '(Cerró)
'adoOrdenes!ord_tipo = 0
'adoOrdenes!ORD_Func = vsFunc
'adoOrdenes!ORD_FDia = vsFDia
'adoOrdenes!ORD_FHora = vsFHora
'adoOrdenes.Update


'===============================
Private Sub Command2_Click()
'===============================
MDIingreso.mInfoAdm.Visible = False
MDIingreso.mInfoAdm.Enabled = False
Unload Me
End Sub

Private Sub Command3_Click()
pb.Visible = True
'para verificar, sumo el total de cuotas
Label1.Caption = "Abriendo tbl_Ordenes...."
Label1.Refresh
adoA.Open "select * FROM tbl_Ordenes;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText

Label1.Caption = "Recorriendo..."
Label1.Refresh

Dim nM As Long
Dim mSuma As Double
pb.Min = 0
pb.Max = adoA.RecordCount


'recorre tbl_ordenes y suma SOLO LAS cuotas

adoA.MoveFirst
Do While Not adoA.EOF
    If nM Mod 100 Then
        pb.Value = nM
    End If
    
    'si NO es una orden anulada
    If Not adoA!ord_tipo = 4 Then
            If adoA!ord_NroOrden = 1 Then
                mSuma = mSuma + adoA!ord_cuota
            End If
    End If
    adoA.MoveNext
    nM = nM + 1
Loop
Label1.Caption = "tOTAL:  " & Format(mSuma, "#,#0.00")
adoA.Close
Set adoA = Nothing
pb.Visible = False
End Sub





'===============================
Private Sub Command4_Click()
'===============================
'resume la tbl_pagos en tbl_JI2


Label1.Caption = "Abriendo tbl_Pagos...."
Label1.Refresh
adoA.Open "select * FROM tbl_Pagos ORDER BY pag_Fecha;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText


Label1.Caption = "Creando tabla virtual..."
Label1.Refresh
'crea una tabla virtual
Set adoB.ActiveConnection = adoconn
Set adoB = New ADODB.Recordset
adoB.Fields.Append "fAM", adChar, 6
adoB.Fields.Append "fMes", adChar, 2
adoB.Fields.Append "fAno", adChar, 4

'pagos
adoB.Fields.Append "fPagosCnts", adInteger, 2  'cuantos
adoB.Fields.Append "fPagosVal", adDouble       'valor en pesos

'pagos atrasados
adoB.Fields.Append "fAtrasoCnts", adInteger, 2      'cuantos
adoB.Fields.Append "fAtrasoDias", adInteger, 2      'suma de dias de atraso

adoB.Fields.Append "fAtrasoMenos30", adInteger, 2        'en fecha: cuantos con 30 dias o menos de atraso
adoB.Fields.Append "fAtraso30a60", adInteger, 2        'cuantas con atraso entre 30 y 60
adoB.Fields.Append "fAtraso60a90", adInteger, 2        'cuantos con 60 a 90
adoB.Fields.Append "fAtraso90a120", adInteger, 2       'cuantos con +90
adoB.Fields.Append "fAtraso120a150", adInteger, 2       'cuantos con +90
adoB.Fields.Append "fAtraso150a180", adInteger, 2       'cuantos con +90
adoB.Fields.Append "fAtraso180a360", adInteger, 2       'cuantos con +90
adoB.Fields.Append "fAtraso360a720", adInteger, 2       'cuantos con +90
adoB.Fields.Append "fAtrasoMas720", adInteger, 2       'cuantos con +90

adoB.Fields.Append "fValAtrasoMenos30", adDouble        'valor con atraso entre 30 y 60
adoB.Fields.Append "fValAtraso30a60", adDouble        'valor con atraso entre 30 y 60
adoB.Fields.Append "fValAtraso60a90", adDouble       'valor con 60 a 90
adoB.Fields.Append "fValAtraso90a120", adDouble     'valor con +90
adoB.Fields.Append "fValAtraso120a150", adDouble       'valor con +90
adoB.Fields.Append "fValAtraso150a180", adDouble       'valor con +90
adoB.Fields.Append "fValAtraso180a360", adDouble      'valor con +90
adoB.Fields.Append "fValAtraso360a720", adDouble       'valor con +90
adoB.Fields.Append "fValAtrasoMas720", adDouble       'valor con +90


adoB.Fields.Append "fAtrasoValor", adDouble         'valor da los pagos atrasados


'generacion de recargo
adoB.Fields.Append "fRecargCnts", adInteger, 2  'cuantos
adoB.Fields.Append "fRecargDias", adInteger, 2  'dias de atraso
adoB.Fields.Append "fRecargVal", adDouble       'valor en pesos




adoB.CursorType = adOpenDynamic
adoB.LockType = adLockOptimistic
adoB.Open

'------------------------------------
'borra los registros de la tabla tbl_JI2
'------------------------------------
Label1.Caption = "Limpiando..."
Label1.Refresh
pb.Visible = True
pb.Min = 0
pb.Max = adoA.RecordCount
adoC.Open "delete * FROM tbl_JI2;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
adoC.Open "select * FROM tbl_JI2;", adoconn, adOpenKeyset, adLockOptimistic, adCmdText





'------------------------------------
'recorre la tabla tbl_Pagos y
'guarda un resumen en  ado
'------------------------------------
Label1.Caption = "Recorriendo..."
Label1.Refresh
Dim nM, nf As Long
nf = adoA.RecordCount
pb.Min = 0
pb.Max = nf
nM = 0

'ordena el ado
adoB.Sort = "fam"

adoA.MoveFirst
Do While Not adoA.EOF
    If (nM Mod 999 = 0) Then
        pb.Value = nM
        lblReg.Caption = nf - nM
        lblReg.Refresh
    End If
    'si NO es una orden anulada
    If Not adoA!pag_Motivo = 4 Then             'si no es orden anulada
        ms2TomaVAM (adoA!pag_Fecha)
        If mfNoEncuentraVam Then
            msAgregaReg
        End If
        ms2GuardaDatosEnAdo
        'Debug.Print adoB!fam, adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU
    End If
    adoA.MoveNext
    nM = nM + 1
Loop


   'Set fjMome.DataGrid1.DataSource = adoB
   ' fjMome.Show
   ' MsgBox "Probando"




'------------------------------------
'guarda en tabla
'------------------------------------
Label1.Caption = "Guardando...."
Label1.Refresh
pb.Min = 0
pb.Max = adoB.RecordCount
pb.Value = 0
adoB.MoveFirst
Do While Not adoB.EOF
   If adoB.AbsolutePosition Mod 100 = 0 Then pb.Value = adoB.AbsolutePosition
    'Debug.Print adoB!fam, adoB!fCntsordP, adoB!fValordP, adoB!fCntsordD, adoB!fValordD, adoB!fCntsordU, adoB!fValordU, adoB!fCntsCuoP, adoB!fValCuoP, adoB!fCntsCuoD, adoB!fValCuoD, adoB!fCntsCuoU, adoB!fValCuoU, adoB!fCntsAyuP, adoB!fValAyuP, adoB!fCntsAyuD, adoB!fValAyuD, adoB!fCntsAyuU, adoB!fValAyuU
    ms2GuardaReg
    adoB.MoveNext
Loop
adoC.Update

'------------------------------------
'formatea los valores decimales de la tabla tbl_JI1
'------------------------------------
Label1.Caption = "Formateando...."
Label1.Refresh
pb.Min = 0
pb.Max = adoC.RecordCount
pb.Value = 0
adoC.MoveFirst
Do While Not adoC.EOF
    pb.Value = adoC.AbsolutePosition
    'Formatea los campos doubles
    ms2FormateaReg adoC(4)
    ms2FormateaReg adoC(16)
    ms2FormateaReg adoC(17)
    ms2FormateaReg adoC(18)
    ms2FormateaReg adoC(19)
    ms2FormateaReg adoC(20)
    ms2FormateaReg adoC(21)
    ms2FormateaReg adoC(22)
    ms2FormateaReg adoC(23)
    ms2FormateaReg adoC(24)
    ms2FormateaReg adoC(25)
    ms2FormateaReg adoC(28)
    adoC.MoveNext
Loop




'------------------------------------
'libera
'------------------------------------
Label1.Caption = ""
Label1.Refresh

adoA.Close
Set adoA = Nothing
adoB.Close
Set adoB = Nothing
adoC.Close
Set adoC = Nothing
pb.Visible = False

End Sub









'===============================
Private Sub ms2GuardaReg()
'===============================
    Dim ni As Integer
    adoC.AddNew
    For ni = 0 To 28
        adoC(ni) = adoB(ni)
        'Debug.Print ni, adoB(ni)
    Next
    adoC.Update
End Sub


'===============================
Private Sub ms2GuardaDatosEnAdo()
'==============================
    Select Case adoA!pag_Motivo
        Case 7      'recargos
            msGuardaValores7
        Case 8      'anulaciones de pagos y recargos
            'recargo
            If Left(adoA!pag_det, 1) = "R" Then
                msGuardaValores8A
            'pagos
            Else
                msGuardaValoresElse (-1)
            End If
        Case Else       'pagos 5=cuota 6=entCta
            msGuardaValoresElse (1)
    End Select
   adoB.Update
End Sub


Private Sub msGuardaValores7()
            adoB!fRecargCnts = adoB!fRecargCnts + 1
            adoB!fRecargDias = adoB!fRecargDias + mfCalculaDIasDeDif
            adoB!fRecargVal = adoB!fRecargVal + adoA!pag_Valor
            adoB.Update
End Sub

Private Sub msGuardaValores8A()
            adoB!fRecargCnts = adoB!fRecargCnts - 1
            adoB!fRecargDias = adoB!fRecargDias - mfCalculaDIasDeDif
            adoB!fRecargVal = adoB!fRecargVal + adoA!pag_Valor
            adoB.Update
End Sub


Private Sub msGuardaValoresElse(nPrm As Integer)
            Dim mDias As Long
            mDias = mf2CalculaDIasDeDif
            
            'clasifica las cantidades de cuentas por atraso
            If mDias > 720 Then
                adoB!fAtrasoMas720 = adoB!fAtrasoMas720 + nPrm
                adoB!fValAtrasoMas720 = adoB!fValAtrasoMas720 + adoA!pag_Valor
           ElseIf mDias > 360 Then
                adoB!fAtraso360a720 = adoB!fAtraso360a720 + nPrm
                adoB!fValAtraso360a720 = adoB!fValAtraso360a720 + adoA!pag_Valor
            ElseIf mDias > 180 Then
                adoB!fAtraso180a360 = adoB!fAtraso180a360 + nPrm
                adoB!fValAtraso180a360 = adoB!fValAtraso180a360 + adoA!pag_Valor
            ElseIf mDias > 150 Then
                adoB!fAtraso150a180 = adoB!fAtraso150a180 + nPrm
                adoB!fValAtraso150a180 = adoB!fValAtraso150a180 + adoA!pag_Valor
            ElseIf mDias > 120 Then
                 adoB!fAtraso120a150 = adoB!fAtraso120a150 + nPrm
                adoB!fValAtraso120a150 = adoB!fValAtraso120a150 + adoA!pag_Valor
          ElseIf mDias > 90 Then
                   adoB!fAtraso90a120 = adoB!fAtraso90a120 + nPrm
                adoB!fValAtraso90a120 = adoB!fValAtraso90a120 + adoA!pag_Valor
          ElseIf mDias > 60 Then
                 adoB!fAtraso60a90 = adoB!fAtraso60a90 + nPrm
                adoB!fValAtraso60a90 = adoB!fValAtraso60a90 + adoA!pag_Valor
           ElseIf mDias > 30 Then
                  adoB!fAtraso30a60 = adoB!fAtraso30a60 + nPrm
                adoB!fValAtraso30a60 = adoB!fValAtraso30a60 + adoA!pag_Valor
           Else
                   adoB!fAtrasoMenos30 = adoB!fAtrasoMenos30 + nPrm
                adoB!fValAtrasoMenos30 = adoB!fValAtrasoMenos30 + adoA!pag_Valor
          End If
            
            'pro los dias de atrsao
            adoB!fAtrasoDias = adoB!fAtrasoDias + mDias * nPrm
            
            'por el valor del pago
            adoB!fPagosVal = adoB!fPagosVal + adoA!pag_Valor
            adoB.Update
End Sub


'===============================
Private Function mfCalculaDIasDeDif() As Long
'===============================
    Dim mDias As Long

    Debug.Print adoA!pag_auto, Mid(adoA!pag_det, 13, 10), adoA!pag_Fecha
    mDias = adoA!pag_Fecha - CDate(Mid(adoA!pag_det, 13, 10))
    Debug.Print mDias, Mid(adoA!pag_det, 13, 10), adoA!pag_Fecha
    mfCalculaDIasDeDif = mDias
End Function



'===============================
Private Function mf2CalculaDIasDeDif() As Long
'===============================
    Dim mDias As Long
    
    'en pag_det estan cuota/plan  fecha de vencimiento
    Debug.Print adoA!pag_auto
    mDias = adoA!pag_Fecha - CDate(Mid(adoA!pag_det, 9, 10))
    Debug.Print mDias, Mid(adoA!pag_det, 9, 10), adoA!pag_Fecha
    
    'para afinar, le saco los dias cuando  no completan un mes
    'mDias = mDias - mDias Mod 30
    If mDias < 30 Then mDias = 0
    Debug.Print mDias
    
    mf2CalculaDIasDeDif = mDias
End Function



'===============================
Private Sub ms2FormateaReg(nPrm As Field)
'===============================
    'deja con 2 decimales unicamente
    
    Dim mS As String
    Dim nc As Integer
    'solo los valores doubles
        mS = nPrm
        mS = mfCambiaPuntoPorComa(mS)
        mS = Format(mS, "0.00")
        nPrm = CDbl(mS)
    adoC.Update

End Sub
'===============================
Private Sub ms2TomaVAM(nPrm As Date)
'===============================
'Toma la fecha de emisión
'si dia > 10 entonces mes++, si mes=13 entonces mes=1;año++
    
    Dim mDia, mMes As Byte
    Dim mAnio As Integer
    
    mDia = Day(nPrm)
    mMes = Month(nPrm)
    mAnio = Year(nPrm)
    
    If mMes < 10 Then
            vAM = mAnio & "0" & mMes
    Else
            vAM = mAnio & mMes
    End If

End Sub
