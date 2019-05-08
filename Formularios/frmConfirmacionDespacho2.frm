VERSION 5.00
Begin VB.Form frmConfirmacionDespacho 
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "frmConfirmacionDespacho"
   ScaleHeight     =   2520
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Despachado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmConfirmacionDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 Public Cod_TipDoc As String
 Public Serie As String
 Public Nro_doc As String
 Public Valor As String
 Public oParent As Object
     
Private Sub Command1_Click()
Dim strSQL As String
Dim cadena As String


If Check1.Value = 1 Then
cadena = "S"
Else
cadena = "N"
End If


    strSQL = "EXEC SP_ActualizarFlgDespaExten  '" & Cod_TipDoc & "','" & Serie & "','" & Nro_doc & "','" & cadena & "'"
           
    ExecuteCommandSQL cCONNECT, strSQL
     
     'oParent.var = "Ubicar"
     oParent.Buscar
    
   

    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Valor = "N" Then
Check1.Value = 0
Else
Check1.Value = 1
End If

End Sub
