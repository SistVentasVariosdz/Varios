VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmCompletarImportesLDPDDP 
   Caption         =   "Completar Importes LDP / DDP"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "frmCompletarImportesLDPDDP"
   ScaleHeight     =   4380
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFlete 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtDdp 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtLdp 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtCif 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtFob 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtTransporte 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtDesaduanaje 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   960
      TabIndex        =   11
      Top             =   3600
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmCompletarImportesLDPDDP.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label7 
      Caption         =   "Flete"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Imp DDP"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Imp LDP"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Imp CIF"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Imp FOB"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Transporte en País Destino "
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Desaduanaje"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCompletarImportesLDPDDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strNum_Corre As String
Public oParent As Object


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo Hand
  
Select Case ActionName
  Case "ACEPTAR"
        calcular
        
     If MsgBox("Esta seguro de grabar los datos", vbInformation + vbYesNo, "AVISO") = vbYes Then
         
      If ValidaCampos() Then
      
         StrSql = "EXEC CN_VENTAS_ACTUALIZAR_PRECIOS_LDP_DDP '" & strNum_Corre & "'," & txtDesaduanaje.Text & "," & txtTransporte.Text & ", " & txtFob.Text & ", " & txtCif.Text & "," & txtLdp.Text & ", " & txtDdp.Text & "," & txtFlete.Text & ""
         Call ExecuteSQL(cConnect, StrSql)
         Unload Me
         
      End If
        
    End If
   
  Case "CANCELAR"
      Unload Me
      
End Select

Exit Sub

Hand:

errores Err.Number
End Sub

 
Private Function ValidaCampos() As Boolean
Dim strMsg As String
 strMsg = ""

If Conversion.CDbl(txtCif.Text) < 0 Then
 strMsg = "El valor CIF no puede ser Negativo"
End If

If Conversion.CDbl(txtLdp.Text) < 0 Then
strMsg = "El valor LDP no puede ser Negativo"
End If

If Conversion.CDbl(txtDdp.Text) < 0 Then
strMsg = "El valor DDP no puede ser Negativo"
End If

 If strMsg <> "" Then
  MsgBox strMsg, vbInformation, "AVISO"
  ValidaCampos = False
  Exit Function
End If

ValidaCampos = True

End Function

 
Private Sub txtCif_GotFocus()
SelectionText txtCif
End Sub

Private Sub txtCif_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtTransporte.SetFocus
End If
End Sub
 

Private Sub txtDdp_GotFocus()
SelectionText txtDdp
End Sub

Private Sub txtDdp_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
    FunctButt1.SetFocus
End If
End Sub
 
 
Private Sub txtDesaduanaje_GotFocus()
SelectionText txtDesaduanaje
End Sub

Private Sub txtDesaduanaje_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtTransporte.SetFocus
End If
End Sub

Private Sub txtFlete_GotFocus()
SelectionText txtFlete
End Sub

Private Sub txtFlete_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    
    txtDesaduanaje.SetFocus
End If
End Sub

Private Sub txtFob_GotFocus()
SelectionText txtFob
End Sub

Private Sub txtFob_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    calcular
    FunctButt1.SetFocus
End If
End Sub



 
Sub calcular()
'Imp_CIF = IMP_FOB + imp_Flete
'Imp_LDP = Imp_CIF + Imp_Desaduanaje
'Imp_DDP = Imp_LDP + Imp_Transporte_Pais_Destino
 
 txtCif.Text = Conversion.CDbl(txtFob.Text) + Conversion.CDbl(txtFlete.Text)
 txtLdp.Text = Conversion.CDbl(txtCif.Text) + Conversion.CDbl(txtDesaduanaje.Text)
 txtDdp.Text = Conversion.CDbl(txtLdp.Text) + Conversion.CDbl(txtTransporte.Text)
 
 
End Sub

 
Private Sub txtLdp_GotFocus()
SelectionText txtLdp
End Sub

Private Sub txtLdp_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    calcular
    FunctButt1.SetFocus
End If
End Sub

Private Sub txtTransporte_GotFocus()
SelectionText txtTransporte
End Sub

Private Sub txtTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtFob.SetFocus
End If
End Sub
