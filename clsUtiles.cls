VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=====================================================================
Public Function DevMesNumerico(sMes As String) As Byte
'=====================================================================
Select Case sMes
    Case "Enero"
        DevMesNumerico = 1
    Case "Febrero"
        DevMesNumerico = 2
    Case "Marzo"
        DevMesNumerico = 3
    Case "Abril"
        DevMesNumerico = 4
    Case "Mayo"
        DevMesNumerico = 5
    Case "Junio"
        DevMesNumerico = 6
    Case "Julio"
        DevMesNumerico = 7
    Case "Agosto"
        DevMesNumerico = 8
    Case "Setiembre"
        DevMesNumerico = 9
    Case "Octubre"
        DevMesNumerico = 10
    Case "Noviembre"
        DevMesNumerico = 11
    Case "Diciembre"
        DevMesNumerico = 12
    Case Else
       DevMesNumerico = 0
End Select
End Function
'=====================================================================
Public Sub AgregaMesALista(cbMes As ListBox)
'=====================================================================
    cbMes.AddItem "Enero"
    cbMes.AddItem "Febrero"
    cbMes.AddItem "Marzo"
    cbMes.AddItem "Abril"
    cbMes.AddItem "Mayo"
    cbMes.AddItem "Junio"
    cbMes.AddItem "Julio"
    cbMes.AddItem "Agosto"
    cbMes.AddItem "Setiembre"
    cbMes.AddItem "Octubre"
    cbMes.AddItem "Noviembre"
    cbMes.AddItem "Diciembre"
End Sub

