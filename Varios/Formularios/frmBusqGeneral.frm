VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8280
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
      Left            =   7110
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   3855
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
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   3855
      Width           =   1065
   End
   Begin GridEX20.GridEX gexList 
      Height          =   3825
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6747
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
      FormatStyle(1)  =   "frmBusqGeneral.frx":0000
      FormatStyle(2)  =   "frmBusqGeneral.frx":0128
      FormatStyle(3)  =   "frmBusqGeneral.frx":01D8
      FormatStyle(4)  =   "frmBusqGeneral.frx":028C
      FormatStyle(5)  =   "frmBusqGeneral.frx":0364
      FormatStyle(6)  =   "frmBusqGeneral.frx":041C
      FormatStyle(7)  =   "frmBusqGeneral.frx":04FC
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneral.frx":051C
   End
End
Attribute VB_Name = "frmBusqGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public oParent As Object
Public sQuery As String
'Dim Rs_Carga As New ADODB.Recordset
Public codigo As String
Public Descripcion As String, paso As Boolean
Public INDICE_CODIGO_AUXILIAR As Integer
Sub Cargar_Datos()
    On Error GoTo Cargar_DatosErr
    
    Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(sQuery, cConnect)
    Dim C  As Integer
    
    With gexList
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
    MsgBox Err.Description, vbCritical, "Cargar_Datos"
End Sub

Private Sub gexList_GroupByBoxHeaderClick(ByVal Group As JSGroup)
    Group.SortOrder = -Group.SortOrder
End Sub

Private Sub gexList_RowFormat(RowBuffer As GridEX20.JSRowData)
    If gexList.RowCount = 0 Then Exit Sub
    Dim fmtConTipoRegistro As JSFmtCondition

    'Set fmtConTipoRegistro = gexList.FmtConditions.Add(gexList.Columns("GKS_CRUDO").Index, jgexEqual, "0.00")

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

Public Sub cmdaceptar_Click()
    gexList_DblClick
End Sub

Public Sub cmdcancelar_Click()
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

Private Sub gexList_DblClick()
On Error Resume Next
gexList.Update
gexList.Refresh
If gexList.RowCount > 0 Then
    If gexList.IsGroupItem(gexList.Row) = True Then Exit Sub
    With oParent
        .codigo = gexList.Value(gexList.Columns(1).Index)
        .Descripcion = gexList.Value(gexList.Columns(2).Index)
        
        'If oParent.Name = "FrmDetalleHilTel" Then
        '    .Cod_color = gexList.Value(gexList.Columns("Cod Color").Index)
        'End If
        .paso = True
        
    End With
    
    If gexList.Columns.Count >= 3 Then
        oParent.CODIGO_AUXILIAR = gexList.Value(gexList.Columns(INDICE_CODIGO_AUXILIAR).Index)
    End If
    With oParent
        .fila_seleccionada = gexList.RowIndex(gexList.Row)
    End With
    
    
'    If gexList.Columns.Count > 3 Then
'        If oParent.Name = "FrmDetalleTelaCa" Then
'            With oParent
'                .Cod_Comb = IIf(IsNull(gexList.Value(gexList.Columns("combinacion").Index)), "", Left(gexList.Value(gexList.Columns("combinacion").Index), 3))
'                .Cod_color = Left(gexList.Value(gexList.Columns("color").Index), 6)
'                .Cod_Talla = gexList.Value(gexList.Columns("talla").Index)
'            End With
'        ElseIf oParent.Name = "FrmDetalleTelCru" Then
'            With oParent
'                .Cod_Comb = IIf(IsNull(gexList.Value(gexList.Columns("combinacion").Index)), "", Left(gexList.Value(gexList.Columns("combinacion").Index), 3))
'                '.Cod_color = Left(Rs_Carga("color"), 6)
'                .Cod_Talla = gexList.Value(gexList.Columns("talla").Index)
'                '.Cod_Calidad = Rs_Carga("calidad")
'            End With
'        ElseIf oParent.Name = "frmDetalleCorte" Then
'            With oParent
'                .sCod_comb = gexList.Value(gexList.Columns("cod_comb").Index)
'                .Label3.Caption = gexList.Value(gexList.Columns("des_comb").Index)
'                .sCod_Color = gexList.Value(gexList.Columns("cod_color").Index)
'                .Label2.Caption = gexList.Value(gexList.Columns("cod_calidad").Index)
'                .lblPartida.Caption = gexList.Value(gexList.Columns("partida").Index)
'                .sCod_Talla = gexList.Value(gexList.Columns("cod_medida").Index)
'                .sCod_TipOrdTra = gexList.Value(gexList.Columns("cod_tipordtra").Index)
'                .sCod_OrdTra = gexList.Value(gexList.Columns("cod_ordtra").Index)
'                .txtcantidad.Text = gexList.Value(gexList.Columns("saldo").Index)
'            End With
'    End If
'
'
'    End If
End If
Unload Me
End Sub

Private Sub gexList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then gexList_DblClick
End Sub


