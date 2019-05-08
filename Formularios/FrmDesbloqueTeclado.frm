VERSION 5.00
Begin VB.Form FrmDesbloqueTeclado 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "FrmDesbloqueTeclado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()

  MsgBox "IMPORTANTE !!! Para desbloquear escribir: 123, " & _
         "y presionar Enter", vbInformation
  ' Inicia El hook para el teclado y para el mouse
  Call IniciarHook(Me.hwnd)

  ' maximizado
  Me.WindowState = 2
  Text1 = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' finaliza el Hook
  Call FinalizarHook(Me.hwnd)
End Sub


' textbox para desbloquear
''''''''''''''''''''''''''''''''''''
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Text1 = "123" And KeyAscii = vbKeyReturn Then
        Call FinalizarHook(Me.hwnd) ' finaliza el Hook
    End If
End Sub

