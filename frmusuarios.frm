VERSION 5.00
Begin VB.Form frmusuarios 
   Caption         =   "Usuarios"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3495
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   1665
      Width           =   2100
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   615
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmusuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Text2_Change()
    Text2.PasswordChar = "#"
End Sub
