VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fjMiReporte 
   Caption         =   "Editor"
   ClientHeight    =   3495
   ClientLeft      =   3000
   ClientTop       =   2880
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6570
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"fjMiReporte.frx":0000
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   3000
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivos"
      Begin VB.Menu mnGuardar 
         Caption         =   "&Guardar..."
         Shortcut        =   ^S
      End
      Begin VB.Menu itmImprimir 
         Caption         =   "&Imnprimir..."
      End
      Begin VB.Menu itmExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edicción"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu itmAcerca 
         Caption         =   "&Acerca"
      End
   End
End
Attribute VB_Name = "fjMiReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
AbreReporte
End Sub


Public Sub AgregaMiReporte(sPrm As String, bsLargo As Byte, nPrm As Single, bnLargo As Byte)
If Len(sPrm) > bsLargo Then
    sPrm = Left(sPrm, bsLargo)
Else
    sPrm = sPrm & Space(bsLargo - Len(sPrm))
End If
txtMain.Text = txtMain.Text & sPrm & vbTab & MiReporteMiFormato(nPrm, bnLargo) & vbCrLf
End Sub





Private Function MiReporteMiFormato(nPrm As Single, bnLargo As Byte) As String
Dim sPrm As String

sPrm = Format(nPrm, "#,#0.00")
sPrm = mRepite(bnLargo - Len(sPrm) * 2, " ") & sPrm
MiReporteMiFormato = sPrm
End Function

Private Function mRepite(nPrm As Byte, sPrm As String) As String
Dim nM As Byte
Dim tPrm As String
For nM = 1 To nPrm
    tPrm = tPrm & sPrm
Next
mRepite = tPrm
End Function



Private Sub Form_Resize()
    txtMain.Top = 0
    txtMain.Left = 0
    txtMain.Width = ScaleWidth
    txtMain.Height = ScaleHeight
End Sub


Private Sub itmAcerca_Click()
    Dim msg$
    
    msg$ = "Editor de Texto" & vbCrLf
    msg$ = msg$ & "por Jun Viera" & vbCrLf
    msg$ = msg$ & Chr$(169) & " Rivera, 2003 "

    MsgBox msg$, vbOKOnly, "Acerca de"

End Sub

Private Sub itmExit_Click()
    Dim msg$
    
    msg$ = "Seguro que desea salir?"
    
    If MsgBox(msg$, vbYesNo + vbQuestion, _
            "Cerrando") = vbYes Then
        Unload Me
    End If
End Sub



Private Sub itmImprimir_Click()
cdMain.Flags = cdlPDReturnDC + cdlPDNoPageNums
   If txtMain.SelLength = 0 Then
      cdMain.Flags = cdMain.Flags + cdlPDAllPages
   Else
      cdMain.Flags = cdMain.Flags + cdlPDSelection
   End If
   cdMain.ShowPrinter
   txtMain.SelPrint cdMain.hDC
End Sub




Private Sub mnGuardar_Click()
    Dim strNombreArchivo As String   'String of file to open
    Dim strFiltro As String     'Common Dialog filter string
    
    'Set the Common Dialog filter
    strFiltro = "Rtf (*.rtf)|*.rtf|Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cdMain.Filter = strFiltro
    
    'Open the common dialog in save mode
    cdMain.ShowSave
    
    'Make sure the retrieved filename is not a blank string
    If cdMain.FileName <> "" Then
        'If it is not blank open the file
        strNombreArchivo = cdMain.FileName
        
       
       
             
        'Set an hour glass cursor just in case it takes a while
        MousePointer = vbHourglass
        
        txtMain.SaveFile strNombreArchivo
        'Reset the cursor to the Windows default.
        MousePointer = vbDefault
        
        End If
    
End Sub






Private Sub mnuCopy_Click()
    Clipboard.SetText txtMain.SelText
End Sub

'no se utiliza
Public Sub Titulo(sPrm As String)
    txtMain.Font.Size = 16
    txtMain.Font.Name = "arial"

    txtMain.Text = sPrm & vbCrLf & vbCrLf
    'txtMain.SelStart = 0
    'txtMain.SelLength = Len(txtMain.Text)
    'txtMain.SelBold = True
    'txtMain.SelFontSize = 16
    'txtMain.SelAlignment = rtfCenter
    'txtMain.SelStart = Len(txtMain.Text)
    
    txtMain.Font.Size = 12
    txtMain.Font.Name = "arial"
    
End Sub
'no se utiliza
Private Sub AbreReporte()
        Dim strText As String       'Contents of file
        Dim strBuffer As String
        'Open "F:\Juan\Proy\Varios\Pruebas\vb NotePad\probando.txt" For Input As #1
        Open "CPTexto.tmp" For Input As #1
        MousePointer = vbHourglass
        Do While Not EOF(1)
            Line Input #1, strBuffer
            strText = strText & strBuffer & vbCrLf
        Loop
        MousePointer = vbDefault
        Close #1
        
        txtMain.Text = strText
        
End Sub

