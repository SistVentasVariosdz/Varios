VERSION 5.00
Begin VB.Form FrmAsignaPartida 
   Caption         =   "Asigna Partida de Stocks para Despacho Exportacion"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdAsignarPartida 
      Caption         =   "Asignar Partida"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtPartidaStocks 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Partida:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmAsignaPartida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StrSQL As String
Public scod_Cliente As String
Public sser_ordcomp As String
Public scod_ordcomp As String
Public ssec_ordcomp As String

Private Sub cmdAsignarPartida_Click()
 If Trim(txtPartidaStocks.Text) = "" Then Exit Sub
 
 If Len(Trim(txtPartidaStocks.Text)) <> 5 Then
    MsgBox "LA partida debe Tener 5 Caracteres", vbInformation + vbOKOnly, "IMPORTANTE"
     Exit Sub
 End If
 
 Call ASIGNAPARTIDASTK
End Sub
Sub ASIGNAPARTIDASTK()
On Error GoTo fin

    StrSQL = "TI_ASIGNA_PARTIDAS_STOCKS '" & Trim(txtPartidaStocks.Text) & "','" & scod_Cliente & "','" & sser_ordcomp & "','" & scod_ordcomp & "','" & ssec_ordcomp & "'"
    Call ExecuteCommandSQL(cConnect, StrSQL)
    MsgBox "se asigno la partida a la orden de servicio", vbInformation + vbOKOnly, "IMPORTANTE"
    Unload Me
Exit Sub
fin:
MsgBox "Ocurrio un Error al asignar la partida" + err.Description, vbInformation + vbOKOnly, "mensaje"

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub txtPartidaStocks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAsignarPartida.SetFocus
    End If
End Sub

Private Sub txtPartidaStocks_LostFocus()
    txtPartidaStocks.Text = Format(txtPartidaStocks.Text, "00000")
End Sub
