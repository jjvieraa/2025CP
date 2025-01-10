Attribute VB_Name = "miBibliot01"
Option Explicit
Public Function mfFormat(sprm As String, _
    fPrm As String) As String
'sPrm es el formato tipo XX/XX/XXXX (mantiene lo que es XX)
'fPrm es la frase a formatear
Dim sM0 As String
Dim sM1 As String
Dim sM2 As String
Dim nM As Byte
Dim nP As Byte
Dim nQ As Byte
'coloca la frase con el formato
nP = Len(fPrm)
For nM = 1 To nP
    sM1 = Mid(fPrm, nM, 1)
    sM2 = Mid(sprm, nM, 1)
    If sM2 = "X" Then
        sM0 = sM0 & sM1
    Else
        sM0 = sM0 & sM2 & sM1
    End If
Next
'coloca el resto del formato
'For nM = nP To Len(sPrm)
'    sM1 = Mid(sPrm, nM, 1)
'    If sM1 = "X" Then
'        sM0 = sM0 + " "
'    Else
'        sM0 = sM0 + sM1
'    End If
'Next


mfFormat = sM0


End Function

Public Function mfCompleta(sprm As String, nPrm As Byte) As String
Dim bb As Byte
Dim sTr As String

bb = Len(sprm)

If Not bb < nPrm Then
    mfCompleta = Left(sprm, nPrm)
    Exit Function
End If
sTr = sprm & Space(nPrm - bb)
mfCompleta = sTr
End Function

Public Sub mAviso(cPrm As String)
    fjAviso.mMensj.Caption = cPrm
    fjAviso.Show vbModal
End Sub
Public Sub mfAviso(sprm As String)
    fjAviso2.Label1.Caption = sprm
    fjAviso2.Show vbModal
End Sub

Public Function mfInvierteMes(ptFecha As String) As String
Dim mpFecha As Date

mpFecha = CDate(ptFecha)
mfInvierteMes = Month(mpFecha) & "/" & _
    Day(mpFecha) & "/" & Year(mpFecha)
End Function


Public Sub mMsgErr(sprm As String)
MsgBox (sprm)
End Sub


Public Function mfEsNulo(sprm As String)
    If IsNull(sprm) Then
        mfEsNulo = ""
    Else
        mfEsNulo = sprm
    End If
End Function
Public Function mfAumentaFechaEnUnMes(fPrm As Date) As Date
Dim bDia As Byte
Dim bMes As Byte
Dim nAnio As Integer


If Not IsDate(fPrm) Then
    mfAumentaFechaEnUnMes = 0
    Exit Function
End If

bDia = Day(fPrm)
bMes = Month(fPrm) + 1
nAnio = Year(fPrm)
Do While Not IsDate(bDia & "/" & bMes & "/" & nAnio)
    bDia = bDia - 1
    If bDia < 28 Then
        MsgBox "Incrementando Mes: Error en Fecha"
        mfAumentaFechaEnUnMes = 0
        Exit Function
    End If
Loop

mfAumentaFechaEnUnMes = CDate(bDia & "/" & bMes & "/" & nAnio)

End Function
Public Function mfArreglaDiaEnFecha(sprm As String) As Date
Dim bDia As Byte
Dim bMes As Byte
Dim nAnio As Integer
Dim fPrm As Date
Dim bLargo As Byte

If IsDate(sprm) Then
    fPrm = CDate(sprm)
    bDia = Day(fPrm)
    bMes = Month(fPrm)
    nAnio = Year(fPrm)
Else
    bLargo = Len(sprm)
    If Mid(sprm, 2, 1) = "/" Then
        bDia = CByte(Left(sprm, 1))
        sprm = "1" & Right(sprm, bLargo - 1)    'momentanem pone dia 1
    Else
        bDia = CByte(Left(sprm, 2))
        sprm = "1" & Right(sprm, bLargo - 2)
    End If
    fPrm = CDate(sprm)
    bMes = Month(fPrm)
    nAnio = Year(fPrm)
End If
Do While Not IsDate(bDia & "/" & bMes & "/" & nAnio)
    bDia = bDia - 1
    If bDia < 28 Then
        MsgBox "Corrigiendo Dia: ERROR en Fecha"
        mfArreglaDiaEnFecha = 0
        Exit Function
    End If
Loop

mfArreglaDiaEnFecha = CDate(bDia & "/" & bMes & "/" & nAnio)
End Function

Public Function mfPalabras(ByVal nPrm As Double, sprm As String) As String
Dim cNstr As String
Dim nMill, nMils, nCien, nDeci As Integer
Dim cMome As String
Dim cFrase As String
Dim sMoneda As String

Select Case sprm
    Case "P"
        sMoneda = "Pesos Uruguayos"
    Case "R"
        sMoneda = "Reales"
    Case "A"
        sMoneda = "Pesos Argentinos"
    Case "D"
        sMoneda = "Dólares"
    Case "R"
        sMoneda = "Unidades Reaj."
    Case Else
        sMoneda = "Desconocida"
End Select

cFrase = ""
If nPrm = 0# Then
    mfPalabras = "Son " & sMoneda & " cero"
   Exit Function
End If
cNstr = Format(nPrm, "000000000000.00")
If nPrm < 0 Then
    cNstr = Space(16 - Len(cNstr)) + cNstr
Else
    cNstr = Space(15 - Len(cNstr)) + cNstr
End If
'MsgBox ("(" + cNstr + ")")
cMome = Mid(cNstr, 1, 6)
nMill = Val(cMome)
cMome = Mid(cNstr, 7, 3)
nMils = Val(cMome)
cMome = Mid(cNstr, 10, 3)
nCien = Val(cMome)
cMome = Mid(cNstr, 14, 2)
nDeci = Val(cMome)
'MsgBox (Str(nMill) + "/" + Str(nMils) + "/" + Str(nCien) + "/" + Str(nDeci) + "/" + CStr(nPrm))
If nMill <> 0 Then
    cFrase = cFrase + mfPlabra1(nMill)
    If nMill = 1 Then
        cFrase = cFrase + "millón "
    Else
        cFrase = cFrase + "millones "
    End If
End If
If nMils <> 0 Then
    If nMils <> 1 Then
        cFrase = cFrase + mfPlabra1(nMils)
    End If
    cFrase = cFrase + "mil "
End If
cFrase = cFrase + mfPlabra1(nCien)
If nDeci <> 0 Then
    cFrase = cFrase + "con " + Format(nDeci, "##") + "/100.-"
Else
    cFrase = cFrase + ".-"
End If
If nPrm > 0 Then
    mfPalabras = "Son " & sMoneda & " " & cFrase
Else
    mfPalabras = "Son " & sMoneda & " (" + cFrase + ")"
End If
'MsgBox (Str(nPrm) + cFrase)
End Function
Function mfPlabra1(ByVal nPrm As Integer) As String
Dim c1, c2, c3 As Integer
Dim cTexto As String
c1 = Int(nPrm / 100)
c2 = Int((nPrm - c1 * 100) / 10)
c3 = nPrm - c1 * 100 - c2 * 10
Select Case c1
    Case 1
        If c2 = 0 And c3 = 0 Then
            cTexto = cTexto + "cien "
        Else
            cTexto = cTexto + "ciento "
        End If
    Case 2
        cTexto = cTexto + "docientos "
    Case 3
        cTexto = cTexto + "trecientos "
    Case 4
        cTexto = cTexto + "cuatrocientos "
    Case 5
        cTexto = cTexto + "quinientos "
    Case 6
        cTexto = cTexto + "seiscientos "
    Case 7
        cTexto = cTexto + "setecientos "
    Case 8
        cTexto = cTexto + "ochocientos "
    Case 9
        cTexto = cTexto + "novecientos "
End Select
Select Case c2
    Case 1
        Select Case c3
            Case 0
                cTexto = cTexto + "diez "
            Case 1
                cTexto = cTexto + "once "
            Case 2
                cTexto = cTexto + "doce "
            Case 3
                cTexto = cTexto + "trece "
            Case 4
                cTexto = cTexto + "catorce "
            Case 5
                cTexto = cTexto + "quince "
            Case 6
                cTexto = cTexto + "dieciseis "
            Case 7
                cTexto = cTexto + "diecisiete "
            Case 8
                cTexto = cTexto + "dieciocho "
            Case 9
                cTexto = cTexto + "diecinueve "
        End Select
    Case 2
        If c3 = 0 Then
            cTexto = cTexto + "veinte "
        Else
            cTexto = cTexto + "veinti "
        End If
    Case 3
        If c3 = 0 Then
            cTexto = cTexto + "treinta "
        Else
            cTexto = cTexto + "treinta y "
        End If
    Case 4
        If c3 = 0 Then
            cTexto = cTexto + "cuarenta "
        Else
            cTexto = cTexto + "cuarenta y "
        End If
    Case 5
        If c3 = 0 Then
            cTexto = cTexto + "cincuenta "
        Else
            cTexto = cTexto + "cincuenta y "
        End If
    Case 6
        If c3 = 0 Then
            cTexto = cTexto + "sesenta "
        Else
            cTexto = cTexto + "sesenta y "
        End If
    Case 7
        If c3 = 0 Then
            cTexto = cTexto + "setenta "
        Else
            cTexto = cTexto + "setenta y "
        End If
    Case 8
        If c3 = 0 Then
            cTexto = cTexto + "ochenta "
        Else
            cTexto = cTexto + "ochenta y "
        End If
    Case 9
        If c3 = 0 Then
            cTexto = cTexto + "noventa "
        Else
            cTexto = cTexto + "noventa y "
        End If
End Select
If c2 <> 1 Then
    Select Case c3
        Case 1
            cTexto = cTexto + "uno "
        Case 2
            cTexto = cTexto + "dos "
        Case 3
            cTexto = cTexto + "tres "
        Case 4
            cTexto = cTexto + "cuatro "
        Case 5
            cTexto = cTexto + "cinco "
        Case 6
            cTexto = cTexto + "seis "
        Case 7
            cTexto = cTexto + "siete "
        Case 8
            cTexto = cTexto + "ocho "
        Case 9
            cTexto = cTexto + "nueve "
End Select
End If
mfPlabra1 = cTexto
End Function

Public Function mfDevMes(nPrm As Byte) As String
Dim sMes As String
    Select Case nPrm
        Case 1
            sMes = "Ene"
        Case 2
            sMes = "Feb"
        Case 3
            sMes = "Mar"
        Case 4
            sMes = "Abr"
        Case 5
            sMes = "May"
        Case 6
            sMes = "Jun"
        Case 7
            sMes = "Jul"
        Case 8
            sMes = "Ago"
        Case 9
            sMes = "Set"
        Case 10
            sMes = "Oct"
        Case 11
            sMes = "Nov"
        Case 12
            sMes = "Dic"
        Case Else
            sMes = "Err"
    End Select
    mfDevMes = sMes
End Function
Public Sub mMsj(mFrs As String)
    MsgBox mFrs
End Sub

Public Function mfEstaVacio(sprm As String) As Boolean
If IsNull(sprm) Or sprm = "" Or Len(sprm) = 0 Then
    mfEstaVacio = True
Else
    mfEstaVacio = False
End If
End Function


Public Function mfRepite(cPrm As String, nPrm As Integer)
Dim nM As Integer
Dim sM As String
For nM = 1 To nPrm
    sM = sM & cPrm
Next nM
mfRepite = sM
End Function

Public Sub msRegistraUnaAccion(bA, sN, sd, sF, sDi, sH)
On Error GoTo xa01
' 4 anula orden
' 8 anula cobro
'21  Prepago
'22      Disq Jef
'23      Disq CP
'24      pago Jef
'25      pago CP
' 26 cierre comercio
' 30 cambio parametros

Dim miComando As ADODB.Command
Set miComando = New ADODB.Command
miComando.ActiveConnection = adoConn
miComando.CommandType = adCmdText
miComando.CommandText = "insert into tbl_Accion values(" & _
            bA & _
            ", '" & Trim(sN) & "'" & _
            ", '" & Trim(sd) & "'" & _
            ", '" & Trim(sF) & "'" & _
            ", '" & Trim(sDi) & "'" & _
            ", '" & Trim(sH) & "', '', '',0)"     'AGREGUE 2 CAMPOS AL FINAL
miComando.Execute
Set miComando = Nothing
Exit Sub
xa01:
MsgBox "Error xa01: Registrando Acción " & Err.Description
End Sub

Public Function mfAgregaMesesAFecha(dFecha As Date, nPrm As Integer) As Date
                   Dim xxdAhora As Date
                    Dim xxnMes As Integer
                    Dim xxnAño As Integer
                    Dim xxnDia As Integer
                    Dim xxnCantMeses As Integer
                
                xxdAhora = dFecha
                xxnMes = Month(xxdAhora)
                xxnAño = Year(xxdAhora)
                xxnDia = Day(xxdAhora)
                xxnMes = xxnMes + nPrm
                Do While xxnMes > 12 Or xxnMes < 1
                    If xxnMes > 12 Then
                            xxnMes = xxnMes - 12
                            xxnAño = xxnAño + 1
                    End If
                    If xxnMes < 1 Then
                        xxnMes = xxnMes + 12
                        xxnAño = xxnAño - 1
                    End If
                Loop
                mfAgregaMesesAFecha = CDate(xxnDia & "/" & xxnMes & "/" & xxnAño)
End Function


Public Function mfCambiaPuntoPorComa(sprm As String) As String
Dim nM As Integer
Dim nT As Integer
Dim sM2 As String
Dim sM3 As String


nM = Len(Trim(sprm))
For nT = 1 To nM
    sM3 = Mid(sprm, nT, 1)
    If sM3 = "." Then
        sM2 = sM2 & ","
    Else
        sM2 = sM2 & sM3
    End If
Next

mfCambiaPuntoPorComa = sM2
End Function

