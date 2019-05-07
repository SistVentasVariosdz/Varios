VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form Frm_Procesos_Tinto_Ex 
   Caption         =   "Procesos De Tintoreria"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFam 
      Height          =   4335
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1560
      TabIndex        =   1
      Top             =   4440
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_Procesos_Tinto_Ex.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "Frm_Procesos_Tinto_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FillProcesos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Cancel = 200
        Me.Hide
    End If

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
Unload Me

End Select
End Sub

Private Sub FillProcesos()
On Error GoTo Fin
Dim sTit As String
    sTit = "Procesos Tintoreria"
   
    strSQL = "select Cod_Proceso_Tinto,Descripcion from Ti_Procesos_Tintoreria"
                 
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    lstFam.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            lstFam.AddItem !Cod_Proceso_Tinto & " " & !descripcion
            .MoveNext
        Loop
        .Close
    End With
    If lstFam.ListCount > 0 Then lstFam.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub


