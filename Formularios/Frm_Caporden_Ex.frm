VERSION 5.00
Begin VB.Form Frm_Caporden_Ex 
   Caption         =   "Captura Orden De Compra"
   ClientHeight    =   1545
   ClientLeft      =   7650
   ClientTop       =   5775
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   3690
   Begin VB.CommandButton cmd_Salir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Cmd_Capturar 
      Caption         =   "&Capturar"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Txt_Numero 
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txt_Serie 
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Orden De Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "Frm_Caporden_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sTipo As String

Private Sub Cmd_Capturar_Click()
    capturar
End Sub

Private Sub Cmd_Salir_Click()
    Unload Me
End Sub

Sub capturar()
Dim i As Integer
Dim sCod_Cliente As String
On Error GoTo hand
    If sTipo = "1" Then
        StrSQL = "EXEC TI_CAPTURA_ORDEN_COMPRA_CONFECCIONES_AUTOMATICA '" & txt_Serie & _
        "','" & Txt_Numero & "'"
    Else
        StrSQL = "EXEC TI_CAPTURA_ORDEN_COMPRA_CONFECCIONES_AUTOMATICA_INKA '" & txt_Serie & _
        "','" & Txt_Numero & "'"
    End If
    
    Call ExecuteSQL(cConnect, StrSQL)
        MsgBox "LA Captura se Realizo con exito"
    Unload Me
    
Exit Sub
hand:
    ErrorHandler err, "Capturar Datos"
End Sub

Private Sub Txt_Numero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Txt_Numero.Text = Format(Trim(Txt_Numero.Text), "000000")
        Cmd_Capturar.SetFocus
    End If
End Sub

Private Sub txt_Serie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_Serie.Text = Format(Trim(txt_Serie.Text), "000")
        Txt_Numero.SetFocus
    End If
End Sub


