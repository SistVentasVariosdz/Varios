VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_FecEnvDoc 
   Caption         =   "Fecha Envio Documento"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "frmConfirmacionDespacho"
   ScaleHeight     =   1695
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   413
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   195035137
         CurrentDate     =   39932
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Envio Documento :"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_FecEnvDoc"
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

    If FixNulos(DTPFecha, vbString) = "" Then
        strSQL = "EXEC VN_Actualiza_FechaEnvioDocum  '" & Cod_TipDoc & "','" & Serie & "','" & Nro_doc & "',null"
     
    Else
        strSQL = "EXEC VN_Actualiza_FechaEnvioDocum  '" & Cod_TipDoc & "','" & Serie & "','" & Nro_doc & "','" & DTPFecha.Value & "'"
        
    End If
    
    ExecuteCommandSQL cCONNECT, strSQL
     
     'oParent.var = "Ubicar"
     oParent.Buscar
    
   

    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

