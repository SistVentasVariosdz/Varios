VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBuscaTela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkAllClient 
      Alignment       =   1  'Right Justify
      Caption         =   "Mostrar Todos los Clientes"
      Height          =   240
      Left            =   4020
      TabIndex        =   3
      Top             =   30
      Width           =   2700
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5775
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4530
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   4320
      Width           =   1215
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3900
      Left            =   75
      TabIndex        =   2
      Top             =   330
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   6879
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmBuscaTela.frx":0000
      FormatStyle(2)  =   "frmBuscaTela.frx":0138
      FormatStyle(3)  =   "frmBuscaTela.frx":01E8
      FormatStyle(4)  =   "frmBuscaTela.frx":029C
      FormatStyle(5)  =   "frmBuscaTela.frx":0374
      FormatStyle(6)  =   "frmBuscaTela.frx":042C
      FormatStyle(7)  =   "frmBuscaTela.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBuscaTela.frx":052C
   End
End
Attribute VB_Name = "frmBuscaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sCod_Tela As String
Public sCod_Cliente As String
Public Campo As Integer
Dim SQuery As String
Dim Rs_Carga As New ADODB.Recordset

Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr

If ChkAllClient.Value = 1 Then
    SQuery = "EXEC TI_BUSCA_TX_TELAS '','" & Trim(sCod_Tela) & "'"
Else
    SQuery = "EXEC TI_BUSCA_TX_TELAS '" & sCod_Cliente & "','" & Trim(sCod_Tela) & "'"
End If

Set DGridLista.ADORecordset = CargarRecordSetDesconectado(SQuery, cConnect)

DGridLista.Columns("abr_Cliente").Caption = "Cliente"
DGridLista.Columns("cod_tela").Caption = "Codigo"
DGridLista.Columns("Des_tela").Caption = "Descripcion"

DGridLista.Columns("abr_Cliente").Width = 700
DGridLista.Columns("cod_tela").Width = 1000
DGridLista.Columns("des_tela").Width = 4000

With oParent
    .CODIGO = ""
    .descripcion = ""
End With

Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub ChkAllClient_Click()
    CARGAR_DATOS
End Sub

Private Sub DGridlista_DblClick()
    With oParent
        If DGridLista.RowCount > 0 Then DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
        .CODIGO = DGridLista.Value(DGridLista.Columns("cod_tela").Index)
        
        If DGridLista.Columns.Count > 1 Then
            If IsNull(DGridLista.Value(DGridLista.Columns("des_tela").Index)) Then
                .descripcion = ""
            Else
                .descripcion = DGridLista.Value(DGridLista.Columns("des_tela").Index)
            End If
        End If
      End With

Unload Me
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DGridlista_DblClick
    End If
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DGridlista_DblClick
    Else
        If DGridLista.RowCount > 0 Then
            Dim Rs As New ADODB.Recordset
            Set Rs = DGridLista.ADORecordset
            Rs.MoveFirst
            Call BuscaCampo(Rs, Rs(Campo).Name, UCase(Chr(KeyAscii)))
            DGridLista.MoveToBookmark Rs.Bookmark
            Set Rs = Nothing
        End If
    End If
End Sub

Private Sub Form_Activate()
    DGridLista.SetFocus
End Sub

Private Sub Form_Load()
Call FormSet(Me)
   SetGeneralGridEX DGridLista, 0, 1
   Campo = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub
Private Sub cmdAceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub cmdCancelar_Click()
oParent.CODIGO = ""
Unload Me
End Sub

