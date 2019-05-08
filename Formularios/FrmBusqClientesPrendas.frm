VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmBusqClientesPrendas 
   Caption         =   "Busqueda de Clientes Prendas"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7230
      TabIndex        =   4
      Tag             =   "&Cancel"
      Top             =   4935
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6045
      TabIndex        =   3
      Tag             =   "&OK"
      Top             =   4935
      Width           =   1065
   End
   Begin VB.TextBox txtDescripcion_Cliente 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.TextBox txtRuc_Cliente 
      Height          =   285
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox txtTip_Anex 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "C"
      Top             =   0
      Width           =   375
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   4545
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8017
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmBusqClientesPrendas.frx":0000
      FormatStyle(2)  =   "FrmBusqClientesPrendas.frx":0128
      FormatStyle(3)  =   "FrmBusqClientesPrendas.frx":01D8
      FormatStyle(4)  =   "FrmBusqClientesPrendas.frx":028C
      FormatStyle(5)  =   "FrmBusqClientesPrendas.frx":0364
      FormatStyle(6)  =   "FrmBusqClientesPrendas.frx":041C
      FormatStyle(7)  =   "FrmBusqClientesPrendas.frx":04FC
      ImageCount      =   0
      PrinterProperties=   "FrmBusqClientesPrendas.frx":051C
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion.:"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "R.U.C.:"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "FrmBusqClientesPrendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public oParent As Object
Public SQuery As String
'Dim Rs_Carga As New ADODB.Recordset
Public CODIGO As String
Public Descripcion As String, paso As Boolean
Public INDICE_CODIGO_AUXILIAR As Integer

Sub Cargar_Datos()
    On Error GoTo Cargar_DatosErr
    
    Set Me.DGridLista.ADORecordset = CargarRecordSetDesconectado(SQuery, cConnect)
    Dim C  As Integer
    
    With DGridLista
        For C = 1 To .Columns.Count
            With .Columns(C)
                .HeaderAlignment = jgexAlignCenter
                .Caption = UCase(Trim(.Caption))
            End With
        Next C
        If .Columns.Count = 2 Then
            .Columns(1).Width = 1200
            .Columns(2).Width = 5000
        End If
    End With
    Exit Sub
Cargar_DatosErr:
    MsgBox err.Description, vbCritical, "Cargar_Datos"
End Sub
Private Sub DGridLista_GroupByBoxHeaderClick(ByVal Group As JSGroup)
    Group.SortOrder = -Group.SortOrder
End Sub
Private Sub DGridLista_RowFormat(RowBuffer As GridEX20.JSRowData)
    If DGridLista.RowCount = 0 Then Exit Sub
    Dim fmtConTipoRegistro As JSFmtCondition

    'Set fmtConTipoRegistro = DGridLista.FmtConditions.Add(DGridLista.Columns("GKS_CRUDO").Index, jgexEqual, "0.00")

    'With fmtConTipoRegistro.FormatStyle
     '   .ForeColor = &H8000&
      '  .FontSize = 8
       ' .BackColor = &H80000018 'vbYellow
    'End With
End Sub

'Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then DGridlista_DblClick
'End Sub

'Private Sub DGridlista_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'If DGridLista.RowContaining(y) >= 1 And DGridLista.RowContaining(y) <= Rs_Carga.RecordCount Then
'    DGridLista.Bookmark = DGridLista.RowBookmark(DGridLista.RowContaining(y))
'End If
'End Sub
'Private Sub Form_Load()
'Call FormSet(Me)
'FormateaGrid DGridLista
'DGridLista.Columns(1).Width = 4000
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'
'    Set Rs_Carga = Nothing
'End Sub

Public Sub CmdAceptar_Click()
    DGridlista_DblClick
End Sub

Public Sub cmdCancelar_Click()
    With oParent
        '.CODIGO = ""
        '.DESCRIPCION = ""
        '.PASO = False
    End With
    Unload Me
End Sub
Private Sub Form_Load()
INDICE_CODIGO_AUXILIAR = 3
End Sub
Private Sub DGridlista_DblClick()
On Error Resume Next
If DGridLista.RowCount > 0 Then
    If DGridLista.IsGroupItem(DGridLista.Row) = True Then Exit Sub
    With oParent
        '.codigo = DGridLista.Value(DGridLista.Columns(1).Index)
        '.Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
        .CODIGO = DGridLista.Value(DGridLista.Columns(1).Index)
        .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
'        frmAdicionaDocumVentasPrendas.txtNum_Ruc.Text = DGridLista.Value(DGridLista.Columns(1).Index)
'        frmAdicionaDocumVentasPrendas.txtDes_TipAne.Text = Trim(DGridLista.Value(DGridLista.Columns(2).Index))
'        frmAdicionaDocumVentasPrendas.txtNum_Ruc.Tag = Trim(DGridLista.Value(DGridLista.Columns(4).Index))
'        frmAdicionaDocumVentasPrendas.txtDes_TipAne.Tag = Trim(DGridLista.Value(DGridLista.Columns(5).Index))
'
'        FrmGuiasRemisionPrendas.txtNum_Ruc.Text = DGridLista.Value(DGridLista.Columns(1).Index)
'        FrmGuiasRemisionPrendas.txtDes_TipAne.Text = Trim(DGridLista.Value(DGridLista.Columns(2).Index))
'        FrmGuiasRemisionPrendas.txtNum_Ruc.Tag = Trim(DGridLista.Value(DGridLista.Columns(4).Index))
'        FrmGuiasRemisionPrendas.txtDes_TipAne.Tag = Trim(DGridLista.Value(DGridLista.Columns(5).Index))
'
        .txtNum_ruc.Text = DGridLista.Value(DGridLista.Columns(1).Index)
        .txtDes_TipAne.Text = Trim(DGridLista.Value(DGridLista.Columns(2).Index))
        .txtNum_ruc.Tag = Trim(DGridLista.Value(DGridLista.Columns(4).Index))
        .txtDes_TipAne.Tag = Trim(DGridLista.Value(DGridLista.Columns(5).Index))
        
        '.paso = True
    End With
    
    If DGridLista.Columns.Count >= 3 Then
        oParent.CODIGO_AUXILIAR = DGridLista.Value(DGridLista.Columns(INDICE_CODIGO_AUXILIAR).Index)
    End If
    
    DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
End If

Unload Me
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    DGridlista_DblClick
    End If
End Sub

Private Sub txtDescripcion_Cliente_Change()
 Call Busca_Opcion_AnexoContable("2", "C", Trim(txtRuc_Cliente.Text), Trim(txtDescripcion_Cliente.Text))
End Sub

Private Sub txtDescripcion_Cliente_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
        'Call Busca_Opcion_AnexoContable("2", "C", Trim(txtRuc_Cliente.Text), Trim(txtDescripcion_Cliente.Text))
    'End If
End Sub

Private Sub txtRuc_Cliente_Change()
  Call Busca_Opcion_AnexoContable("1", "C", Trim(txtRuc_Cliente.Text), Trim(txtDescripcion_Cliente.Text))
End Sub

Private Sub txtRuc_Cliente_KeyPress(KeyAscii As Integer)
 'If KeyAscii = 13 Then
   'Call Busca_Opcion_AnexoContable("1", "C", Trim(txtRuc_Cliente.Text), Trim(txtDescripcion_Cliente.Text))
 'End If
End Sub
Public Sub Busca_Opcion_AnexoContable(sTipo As String, txttipo As String, ruc As String, txtDes As String)
On Error GoTo fin

Dim rstAux As Object, strsql As String
Set rstAux = CreateObject("ADODB.Recordset")
    'StrSql = "CN_MUESTRA_ANEXOS_CLIENTES '" & sTipo & "','" & txttipo & "','" & ruc & "','" & txtDes & "'"
    strsql = "CN_MUESTRA_ANEXOS_CLIENTES_PRENDAS '" & sTipo & "','" & txttipo & "','" & ruc & "','" & txtDes & "'"
    
    With FrmBusqClientesPrendas
        .SQuery = strsql
        .Cargar_Datos
        
        CODIGO = ""
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Tipo").Width = 800
        .DGridLista.Columns("Nombre").Width = 4075
        .DGridLista.Columns("RUC").Width = 1200
        Set rstAux = .DGridLista.ADORecordset
    
    End With
Exit Sub
fin:
On Error Resume Next
    Unload FrmBusqClientesPrendas
    Set FrmBusqClientesPrendas = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento "
End Sub
'






