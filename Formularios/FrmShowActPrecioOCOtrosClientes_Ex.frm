VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShow_PrecioOCOtrosClientes_Ex 
   Caption         =   "Ordenes Compra Tintoreria Otros Clientes"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txt_talla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txt_combo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10455
      Begin VB.TextBox Txt_Tela 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   14
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Combo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   6
         Top             =   330
         Width           =   600
      End
      Begin VB.Label LblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label LblSecuencia 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   300
         Width           =   615
      End
      Begin VB.Label LblOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   300
         Width           =   1695
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2640
      TabIndex        =   1
      Top             =   5040
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   900
      Custom          =   $"FrmShowActPrecioOCOtrosClientes_Ex.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   6456
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":013E
      Column(2)       =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":0206
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":02AA
      FormatStyle(2)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":03E2
      FormatStyle(3)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":0492
      FormatStyle(4)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":0546
      FormatStyle(5)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":061E
      FormatStyle(6)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":06D6
      FormatStyle(7)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":07B6
      FormatStyle(8)  =   "FrmShowActPrecioOCOtrosClientes_Ex.frx":0862
      ImageCount      =   0
      PrinterProperties=   "FrmShowActPrecioOCOtrosClientes_Ex.frx":0912
   End
End
Attribute VB_Name = "FrmShow_PrecioOCOtrosClientes_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String, sSer_OrdComp As String, sCod_Ordcomp As String, sSec_Ordcomp As String
Dim strSQL As String

Sub CARGA_GRID()
strSQL = "Ventas_ocs_sin_precio_tejeduria_otros_Clientes '" & sCod_Cliente & "','" & sSer_OrdComp & "','" & sCod_Ordcomp & "','" & sSec_Ordcomp & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

GridEX1.Columns("Otro_Cliente").Width = 2300
GridEX1.Columns("Moneda").Width = 800
GridEX1.Columns("Precio").Width = 900
GridEX1.Columns("porc_igv").Width = 750
GridEX1.Columns("Condicion_Venta").Width = 2000
GridEX1.Columns("Descuento").Width = 1700
GridEX1.Columns("cod_otro_cliente").Width = 0

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim i As Integer
Select Case ActionName
Case "ADICIONAR"
    Load FrmAddOCOtrosCliente
    FrmAddOCOtrosCliente.sAccion = "I"
    FrmAddOCOtrosCliente.sCod_Cliente = Me.sCod_Cliente
    FrmAddOCOtrosCliente.sCod_Ordcomp = Me.sCod_Ordcomp
    FrmAddOCOtrosCliente.sSer_OrdComp = Me.sSer_OrdComp
    FrmAddOCOtrosCliente.sSec_Ordcomp = Me.sSec_Ordcomp
    strSQL = "SELECT Porc_IGV FROM TG_IGV WHERE ANO=YEAR(GETDATE()) AND MES=RIGHT('0'+CONVERT(VARCHAR,MONTH(GETDATE())),2) "
    FrmAddOCOtrosCliente.TxtIGV.Text = DevuelveCampo(strSQL, cConnect)
    FrmAddOCOtrosCliente.Show vbModal
    Set FrmAddOCOtrosCliente = Nothing
    Call CARGA_GRID
Case "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmAddOCOtrosCliente
    FrmAddOCOtrosCliente.sAccion = "U"
    FrmAddOCOtrosCliente.sCod_Cliente = Me.sCod_Cliente
    FrmAddOCOtrosCliente.sCod_Ordcomp = Me.sCod_Ordcomp
    FrmAddOCOtrosCliente.sSer_OrdComp = Me.sSer_OrdComp
    FrmAddOCOtrosCliente.sSec_Ordcomp = Me.sSec_Ordcomp
    FrmAddOCOtrosCliente.TxtCod_Cliente.Enabled = False
    FrmAddOCOtrosCliente.TxtDes_Cliente.Enabled = False
    FrmAddOCOtrosCliente.FraMod.Enabled = True
    FrmAddOCOtrosCliente.TxtCod_Cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 1, 5)
    FrmAddOCOtrosCliente.TxtDes_Cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 6)
    FrmAddOCOtrosCliente.TxtCod_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 1, 3)
    FrmAddOCOtrosCliente.TxtDes_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 4)
    FrmAddOCOtrosCliente.TxtPrecio.Text = CDbl(GridEX1.Value(GridEX1.Columns("precio").Index))
    FrmAddOCOtrosCliente.TxtIGV.Text = CDbl(GridEX1.Value(GridEX1.Columns("porc_igv").Index))
    FrmAddOCOtrosCliente.Txtcod_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 1, 3)
    FrmAddOCOtrosCliente.TxtDes_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 4)
    FrmAddOCOtrosCliente.TxtCod_Moneda.Text = GridEX1.Value(GridEX1.Columns("moneda").Index)
    FrmAddOCOtrosCliente.TxtNom_moneda.Text = DevuelveCampo("select nom_moneda from tg_moneda where cod_moneda='" & GridEX1.Value(GridEX1.Columns("moneda").Index) & "'", cConnect)
    FrmAddOCOtrosCliente.Show vbModal
    Set FrmAddOCOtrosCliente = Nothing
    Call CARGA_GRID
Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmAddOCOtrosCliente
    FrmAddOCOtrosCliente.sAccion = "D"
    FrmAddOCOtrosCliente.sCod_Cliente = Me.sCod_Cliente
    FrmAddOCOtrosCliente.sCod_Ordcomp = Me.sCod_Ordcomp
    FrmAddOCOtrosCliente.sSer_OrdComp = Me.sSer_OrdComp
    FrmAddOCOtrosCliente.sSec_Ordcomp = Me.sSec_Ordcomp
    
    FrmAddOCOtrosCliente.TxtCod_Cliente.Enabled = False
    FrmAddOCOtrosCliente.TxtDes_Cliente.Enabled = False
    FrmAddOCOtrosCliente.FraMod.Enabled = False
    FrmAddOCOtrosCliente.TxtCod_Cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 1, 5)
    FrmAddOCOtrosCliente.TxtDes_Cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 6)
    FrmAddOCOtrosCliente.TxtCod_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 1, 3)
    FrmAddOCOtrosCliente.TxtDes_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 4)
    FrmAddOCOtrosCliente.TxtPrecio.Text = CDbl(GridEX1.Value(GridEX1.Columns("precio").Index))
    FrmAddOCOtrosCliente.TxtIGV.Text = CDbl(GridEX1.Value(GridEX1.Columns("porc_igv").Index))
    FrmAddOCOtrosCliente.Txtcod_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 1, 3)
    FrmAddOCOtrosCliente.TxtDes_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 4)
    FrmAddOCOtrosCliente.TxtCod_Moneda.Text = GridEX1.Value(GridEX1.Columns("moneda").Index)
    FrmAddOCOtrosCliente.TxtNom_moneda.Text = DevuelveCampo("select nom_moneda from tg_moneda where cod_moneda='" & GridEX1.Value(GridEX1.Columns("moneda").Index) & "'", cConnect)
    FrmAddOCOtrosCliente.Show vbModal
    Set FrmAddOCOtrosCliente = Nothing
    Call CARGA_GRID
'Case "ACTUALIZAR"
'    For i = 1 To GridEX1.RowCount
'        GridEX1.Row = i
'        Call Actualizar_Precios
'    Next
'    Call CARGA_GRID
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
If ColIndex <> GridEX1.Columns("Precio").Index Then
    Cancel = True
End If
End Sub

Sub Actualizar_Precios()
On Error GoTo errActualizar

strSQL = "Ventas_Up_Actualiza_Precio_Tejeduria_Otros_Clientes '" & sCod_Cliente & "','" & sSer_OrdComp & "','" & sCod_Ordcomp & "','" & sSec_Ordcomp & "','" & GridEX1.Value(GridEX1.Columns("cod_otro_cliente").Index) & "'," & GridEX1.Value(GridEX1.Columns("Precio").Index)
ExecuteCommandSQL cConnect, strSQL

Exit Sub
errActualizar:
    MsgBox err.Description, vbCritical, "Actualizar Precios"
End Sub

Private Sub GridEX1_GotFocus()
'GridEX1.Col = GridEX1.Columns("Precio").Index
End Sub

Private Sub GridEX1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    GridEX1.Row = GridEX1.Row + 1
End If
End Sub
