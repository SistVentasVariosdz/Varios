VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form Frm_Añadir_Procesos_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtDes_Proceso_Tinto 
         Height          =   285
         Left            =   2475
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   2430
      End
      Begin VB.TextBox txtCod_Proceso_Tinto 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   0
         Top             =   960
         Width           =   750
      End
      Begin VB.TextBox Txt_Observaciones 
         Height          =   855
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Txt_Sec 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Txt_Numero 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Serie 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Txt_Cliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sec"
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesos"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden De Compra"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1275
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_Añadir_Procesos_Ex.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "Frm_Añadir_Procesos_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SCod_Cliente_Tex As String
Public Ser_OrdComp     As String
Public Cod_OrdComp     As String
Public Sec_OrdComp     As String
Public sAccion     As String
Public CODIGO As String
Public Descripcion  As String

Private Sub Txt_Procesos_Change()

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        Grabar
    Case "CANCELAR"
        Unload Me
    
End Select
End Sub

Private Sub txtCod_Proceso_Tinto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaProceso 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_Proceso_Tinto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaProceso 2
        SendKeys "{TAB}"
    End If
End Sub


Sub Grabar()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand
    StrSQL = "EXEC Ti_Up_Procesos_Ordencompra '" & sAccion & "','" & SCod_Cliente_Tex & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Sec_OrdComp & "','" & txtCod_Proceso_Tinto & "','" & Txt_Observaciones & "'"
    Call ExecuteSQL(cConnect, StrSQL)
    Unload Me
    
Exit Sub
hand:
    ErrorHandler err, "Capturar Datos"
End Sub


Public Sub BuscaProceso(Opcion As Integer)
On Error GoTo fin
    
    StrSQL = "SELECT Cod_Proceso_Tinto, Descripcion FROM TI_PROCESOS_TINTORERIA WHERE "
    txtCod_Proceso_Tinto = Trim(txtCod_Proceso_Tinto)
    txtDes_Proceso_Tinto = Trim(txtDes_Proceso_Tinto)
    
    Select Case Opcion
    Case 1: StrSQL = StrSQL & " Cod_Proceso_Tinto LIKE '%" & txtCod_Proceso_Tinto & "%'"
    Case 2: StrSQL = StrSQL & " Descripcion LIKE '%" & txtDes_Proceso_Tinto & "%'"
    End Select
    
    txtCod_Proceso_Tinto = ""
    txtDes_Proceso_Tinto = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        
        .gexList.Columns("Cod_Proceso_Tinto").Caption = "Codigo"
        .gexList.Columns("Cod_Proceso_Tinto").Width = 1000
        .gexList.Columns("Descripcion").Caption = "Proceso"
        .gexList.Columns("Descripcion").Width = 5000
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        CODIGO = ".."
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_Proceso_Tinto = Trim(rstAux!COD_PROCESO_TINTO)
            txtDes_Proceso_Tinto = Trim(rstAux!Descripcion)
        End If
    End With
    Unload frmBusqGeneral
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
    On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
    MsgBox err.Description, vbCritical + vbOKOnly, _
    "Busqueda de Proceso (" & Opcion & ")"
End Sub


