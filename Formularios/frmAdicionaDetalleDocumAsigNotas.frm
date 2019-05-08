VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdicionaDetalleDocumAsigNotas 
   ClientHeight    =   6120
   ClientLeft      =   420
   ClientTop       =   1035
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8715
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   435
      Left            =   7080
      TabIndex        =   4
      Top             =   240
      Width           =   1545
   End
   Begin VB.Frame frNroDoc 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5640
         MaxLength       =   8
         TabIndex        =   3
         Top             =   375
         Width           =   1080
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   2
         Top             =   375
         Width           =   540
      End
      Begin VB.TextBox txtCod_TipDoc 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   0
         Top             =   375
         Width           =   480
      End
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Número :"
         Height          =   195
         Left            =   4920
         TabIndex        =   7
         Tag             =   "Number"
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label12 
         Caption         =   "Serie :"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   390
         Width           =   495
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   2940
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   5186
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0000
      FormatStyle(2)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0138
      FormatStyle(3)  =   "frmAdicionaDetalleDocumAsigNotas.frx":01E8
      FormatStyle(4)  =   "frmAdicionaDetalleDocumAsigNotas.frx":029C
      FormatStyle(5)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0374
      FormatStyle(6)  =   "frmAdicionaDetalleDocumAsigNotas.frx":042C
      FormatStyle(7)  =   "frmAdicionaDetalleDocumAsigNotas.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmAdicionaDetalleDocumAsigNotas.frx":052C
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   1380
      Left            =   0
      TabIndex        =   10
      Top             =   3975
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   2434
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0704
      FormatStyle(2)  =   "frmAdicionaDetalleDocumAsigNotas.frx":083C
      FormatStyle(3)  =   "frmAdicionaDetalleDocumAsigNotas.frx":08EC
      FormatStyle(4)  =   "frmAdicionaDetalleDocumAsigNotas.frx":09A0
      FormatStyle(5)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0A78
      FormatStyle(6)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0B30
      FormatStyle(7)  =   "frmAdicionaDetalleDocumAsigNotas.frx":0C10
      ImageCount      =   0
      PrinterProperties=   "frmAdicionaDetalleDocumAsigNotas.frx":0C30
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2880
      TabIndex        =   11
      Top             =   5520
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDetalleDocumAsigNotas.frx":0E08
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmAdicionaDetalleDocumAsigNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODIGO As String, Descripcion As String
Public strNum_Corre_Ori
Private Sub cmdBuscar_Click()
 BUSCAR
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo dprDepurar

If ActionName = "ACEPTAR" Then
  If GridEX1.RowCount = 0 Or GridEX2.RowCount = 0 Then
    MsgBox "Seleccione un Item de Factura", vbInformation, "AVISO"
    Exit Sub
  End If
End If
  
Load frmAdicionaDetalleDocum
With frmAdicionaDetalleDocum
  .Caption = Me.Caption
  .strNum_Corre_Detalle = strNum_Corre_Ori
  .IntSencuencia = 0
  .StrOption = "I"
  If ActionName = "ACEPTAR" Then
    .strNum_Corre_Doc_Asig = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
    .IntSencuencia_Doc_Asig = GridEX2.Value(GridEX2.Columns("Secuencia").Index)
    .txtTip_Item = "P"
    .txtCod_Producto = GridEX2.Value(GridEX2.Columns("Codigo_Item").Index)
    .TxtDescripcion = GridEX2.Value(GridEX2.Columns("Descripcion").Index)
    .TxtCantidad.Text = GridEX2.Value(GridEX2.Columns("Cantidad").Index)
    .txtImp_Unitario.Text = GridEX2.Value(GridEX2.Columns("Imp_Unitario").Index)
    .txtImp_Total.Text = GridEX2.Value(GridEX2.Columns("Imp_Total").Index)
  End If
  Me.Visible = False
  .Show 1
  Unload Me
End With

Exit Sub

dprDepurar:

errores err.Number

End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)

Dim StrSQL As String

If GridEX1.RowCount <> 0 Then
  StrSQL = "Ventas_Muestras_Facturas_por_Asignar_Item '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"
  Set GridEX2.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
  GridEX2.Columns("Codigo_Item").Width = 1500
  GridEX2.Columns("Descripcion").Width = 4530
  GridEX2.Columns("Cantidad").Width = 765
  GridEX2.Columns("Imp_Total").Width = 840
  GridEX2.Columns("Tipo").Width = 450
  GridEX2.Columns("secuencia").Visible = False
End If

End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
End Sub

Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtNum_Docum_LostFocus()
  txtNum_Docum = Format(txtNum_Docum, "00000000")
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Sub BUSCAR()

On Error GoTo dprDepurar

Dim sSQL As String

sSQL = "Ventas_Muestras_Facturas_por_Asignar '" & strNum_Corre_Ori & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Columns("Cod_TipDoc").Width = 495
GridEX1.Columns("Cod_TipDoc").Caption = "Tip"
GridEX1.Columns("Ser_Docum").Width = 525
GridEX1.Columns("Ser_Docum").Caption = "Serie"
GridEX1.Columns("Num_Docum_Ventas").Width = 885
GridEX1.Columns("Num_Docum_Ventas").Caption = "Nro_Doc"
GridEX1.Columns("Fecha").Width = 1200
GridEX1.Columns("Cod_Moneda").Width = 1050
GridEX1.Columns("Ano_Registro").Width = 1080
GridEX1.Columns("Mes_Registro").Width = 1185
GridEX1.Columns("Fecha").Caption = "Emision"
GridEX1.Columns("Fecha").Format = "dd/mm/yyy"
GridEX1.Columns("Num_Corre").Visible = False
GridEX1.Columns("Imp_Total").Width = 1500

GridEX1.ContinuousScroll = True

Exit Sub

dprDepurar:

errores err.Number
  
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub
