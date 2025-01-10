VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fjDatCom 
   Caption         =   "Rubros"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin MSAdodcLib.Adodc adorubro 
         Height          =   330
         Left            =   480
         Top             =   2400
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=jimmy"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "jimmy"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "TBL_Comercios"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtRubro 
         DataField       =   "callnom"
         DataSource      =   "adoCalle"
         Height          =   315
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmDatCom.frx":0000
         Height          =   1875
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3307
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IdRubro"
            Caption         =   "Código "
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
            DataField       =   "Desc"
            Caption         =   "Descripción"
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
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   1440
         TabIndex        =   3
         Top             =   3480
         Width           =   4935
         Begin VB.CommandButton cmdsalir 
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "&Borrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Nuevo Rubro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Existentes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
   End
End
Attribute VB_Name = "fjDatCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rstRubro As Recordset
Attribute rstRubro.VB_VarHelpID = -1
Private Sub cmdGuardar_Click()
'CATEGORIA
On Error GoTo Errores
Dim cod As Integer
Dim codul As Integer
Dim res As Integer
Dim i As Integer
Dim comando As ADODB.Command
Set comando = New ADODB.Command
comando.ActiveConnection = adoconn
comando.CommandType = adCmdText

Set rstRubro = New Recordset
If rstRubro.State = adStateOpen Then rstRubro.Close
rstRubro.Open "select * from Rubro", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstRubro.Sort = "idrubro"
res = vbYes
    If rstRubro.RecordCount <> 0 Then
    'Busca la categoria para verificar que no exista
        rstRubro.MoveFirst
        For i = 1 To rstRubro.RecordCount
            If Trim(rstRubro!Desc) = Trim(txtRubro.Text) Then
                res = MsgBox("El Rubro " & txtRubro.Text & " ya existe en la Base de Datos" & Chr(13) & " ¿ Esta seguro que desea guardar este nuevo registro ?", vbQuestion + vbYesNo, "Pregunta")
                Exit For
            End If
        rstRubro.MoveNext
        Next i
        End If
    
    Select Case res
    Case vbYes
        'Busca el Codigo a asignar
        If rstRubro.RecordCount <> 0 Then
            rstRubro.MoveFirst
            codul = rstRubro!idrubro
            For i = 1 To rstRubro.RecordCount
                rstRubro.MoveNext
                If i <> rstRubro.RecordCount Then
                    If (rstRubro!idrubro - codul) <> 1 Then
                        cod = codul
                        Exit For
                    Else
                        codul = rstRubro!idrubro
                    End If
                Else
                    cod = codul
                    Exit For
                End If
            Next i
        Else
            cod = 0
        End If
        comando.CommandText = "insert into Rubro values(" & Val(cod + 1) & ", '" & Trim(txtRubro.Text) & "')"
        comando.Execute
        MsgBox "El Rubro " & txtRubro.Text & " se ha guardado satisfactoriamente", vbInformation, "Círculo Policial"
        txtRubro.Text = ""
        txtRubro.SetFocus
        adorubro.Refresh
        'Guarda y descarga el formulario
        'Unload Me
    Case vbNo
        MsgBox "No se ha guardado", vbInformation, "Respuesta"
        txtRubro.Text = ""
        txtRubro.SetFocus
    End Select
Exit Sub
Errores:
    MsgBox Err.Description
End Sub
Private Sub cmdBorrar_Click()
'CATEGORIA
Dim Borrar As ADODB.Command
On Error GoTo errorBorrar
Set Borrar = New Command

Set rstRubro = New Recordset
If rstRubro.State = adStateOpen Then rstRubro.Close
rstRubro.Open "select * from rubro", adoconn, adOpenKeyset, adLockOptimistic, adCmdText
rstRubro.Sort = "idrubro"

If DataGrid1.Columns(0).Text <> 0 Then
    Borrar.CommandText = "delete from Rubro where idRubro = " & DataGrid1.Columns(0).Text
    Set Borrar.ActiveConnection = adoconn
    Borrar.Execute
    Borrar.Execute
    adorubro.Refresh
    Exit Sub
End If
Exit Sub

errorBorrar:
    If Err.Number = -2147217900 Then
        MsgBox "No se puede eliminar este registro ya que esta siendo utilizado en la Base de Datos", vbCritical, "Error"
    Else
        MsgBox Err.Description
    End If
End Sub
Private Sub cmdsalir_Click()
Unload Me
End Sub


