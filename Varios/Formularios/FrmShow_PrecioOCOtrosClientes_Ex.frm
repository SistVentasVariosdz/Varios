VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShow_PrecioOCOtrosClientes_Ex 
   Caption         =   "Ordenes Compra Tintoreria Otros Clientes"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
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
      TabIndex        =   14
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
      TabIndex        =   12
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
      TabIndex        =   10
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
      TabIndex        =   1
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
         TabIndex        =   8
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4020
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   7091
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
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":0000
      FormatStyle(2)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":0138
      FormatStyle(3)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":01E8
      FormatStyle(4)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":029C
      FormatStyle(5)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":0374
      FormatStyle(6)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":042C
      FormatStyle(7)  =   "FrmShow_PrecioOCOtrosClientes_Ex.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmShow_PrecioOCOtrosClientes_Ex.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2760
      TabIndex        =   15
      Top             =   5400
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   900
      Custom          =   $"FrmShow_PrecioOCOtrosClientes_Ex.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmShow_PrecioOCOtrosClientes_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String, sSer_OrdComp As String, sCod_OrdComp As String, sSec_OrdComp As String
Dim strSQL As String

Sub CARGA_GRID()
strSQL = "Ventas_ocs_sin_precio_tejeduria_otros_Clientes '" & sCod_Cliente & "','" & sSer_OrdComp & "','" & sCod_OrdComp & "','" & sSec_OrdComp & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

GridEX1.Columns("Otro_Cliente").Width = 2300
GridEX1.Columns("Moneda").Width = 800
GridEX1.Columns("Precio").Width = 900
GridEX1.Columns("porc_igv").Width = 750
GridEX1.Columns("Condicion_Venta").Width = 2000
GridEX1.Columns("Descuento").Width = 1700
GridEX1.Columns("cod_otro_cliente").Width = 0

End Sub

'Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Dim i As Integer
'Select Case ActionName
'Case "ADICIONAR"
'    Load FrmAddOCOtrosCliente
'    FrmAddOCOtrosCliente.saccion = "I"
'    FrmAddOCOtrosCliente.sCod_Cliente = Me.sCod_Cliente
'    FrmAddOCOtrosCliente.sCod_OrdComp = Me.sCod_OrdComp
'    FrmAddOCOtrosCliente.sSer_OrdComp = Me.sSer_OrdComp
'    FrmAddOCOtrosCliente.sSec_OrdComp = Me.sSec_OrdComp
'    strSQL = "SELECT Porc_IGV FROM TG_IGV WHERE ANO=YEAR(GETDATE()) AND MES=RIGHT('0'+CONVERT(VARCHAR,MONTH(GETDATE())),2) "
'    FrmAddOCOtrosCliente.TxtIGV.Text = DevuelveCampo(strSQL, cConnect)
'    FrmAddOCOtrosCliente.Show vbModal
'    Set FrmAddOCOtrosCliente = Nothing
'    Call CARGA_GRID
'Case "MODIFICAR"
'    If GridEX1.RowCount = 0 Then Exit Sub
'    Load FrmAddOCOtrosCliente
'    FrmAddOCOtrosCliente.saccion = "U"
'    FrmAddOCOtrosCliente.sCod_Cliente = Me.sCod_Cliente
'    FrmAddOCOtrosCliente.sCod_OrdComp = Me.sCod_OrdComp
'    FrmAddOCOtrosCliente.sSer_OrdComp = Me.sSer_OrdComp
'    FrmAddOCOtrosCliente.sSec_OrdComp = Me.sSec_OrdComp
'    FrmAddOCOtrosCliente.txtCod_cliente.Enabled = False
'    FrmAddOCOtrosCliente.txtDEs_cliente.Enabled = False
'    FrmAddOCOtrosCliente.FraMod.Enabled = True
'    FrmAddOCOtrosCliente.txtCod_cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 1, 5)
'    FrmAddOCOtrosCliente.txtDEs_cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 6)
'    FrmAddOCOtrosCliente.TxtCod_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 1, 3)
'    FrmAddOCOtrosCliente.TxtDes_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 4)
'    FrmAddOCOtrosCliente.txtPrecio.Text = CDbl(GridEX1.Value(GridEX1.Columns("precio").Index))
'    FrmAddOCOtrosCliente.TxtIGV.Text = CDbl(GridEX1.Value(GridEX1.Columns("porc_igv").Index))
'    FrmAddOCOtrosCliente.txtCod_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 1, 3)
'    FrmAddOCOtrosCliente.txtDes_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 4)
'    FrmAddOCOtrosCliente.txtCod_Moneda.Text = GridEX1.Value(GridEX1.Columns("moneda").Index)
'    FrmAddOCOtrosCliente.TxtNom_moneda.Text = DevuelveCampo("select nom_moneda from tg_moneda where cod_moneda='" & GridEX1.Value(GridEX1.Columns("moneda").Index) & "'", cConnect)
'    FrmAddOCOtrosCliente.Show vbModal
'    Set FrmAddOCOtrosCliente = Nothing
'    Call CARGA_GRID
'Case "ELIMINAR"
'    If GridEX1.RowCount = 0 Then Exit Sub
'    Load FrmAddOCOtrosCliente
'    FrmAddOCOtrosCliente.saccion = "D"
'    FrmAddOCOtrosCliente.sCod_Cliente = Me.sCod_Cliente
'    FrmAddOCOtrosCliente.sCod_OrdComp = Me.sCod_OrdComp
'    FrmAddOCOtrosCliente.sSer_OrdComp = Me.sSer_OrdComp
'    FrmAddOCOtrosCliente.sSec_OrdComp = Me.sSec_OrdComp
'
'    FrmAddOCOtrosCliente.txtCod_cliente.Enabled = False
'    FrmAddOCOtrosCliente.txtDEs_cliente.Enabled = False
'    FrmAddOCOtrosCliente.FraMod.Enabled = False
'    FrmAddOCOtrosCliente.txtCod_cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 1, 5)
'    FrmAddOCOtrosCliente.txtDEs_cliente.Text = Mid(GridEX1.Value(GridEX1.Columns("otro_cliente").Index), 6)
'    FrmAddOCOtrosCliente.TxtCod_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 1, 3)
'    FrmAddOCOtrosCliente.TxtDes_Condicion.Text = Mid(GridEX1.Value(GridEX1.Columns("condicion_venta").Index), 4)
'    FrmAddOCOtrosCliente.txtPrecio.Text = CDbl(GridEX1.Value(GridEX1.Columns("precio").Index))
'    FrmAddOCOtrosCliente.TxtIGV.Text = CDbl(GridEX1.Value(GridEX1.Columns("porc_igv").Index))
'    FrmAddOCOtrosCliente.txtCod_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 1, 3)
'    FrmAddOCOtrosCliente.txtDes_Descuento.Text = Mid(GridEX1.Value(GridEX1.Columns("descuento").Index), 4)
'    FrmAddOCOtrosCliente.txtCod_Moneda.Text = GridEX1.Value(GridEX1.Columns("moneda").Index)
'    FrmAddOCOtrosCliente.TxtNom_moneda.Text = DevuelveCampo("select nom_moneda from tg_moneda where cod_moneda='" & GridEX1.Value(GridEX1.Columns("moneda").Index) & "'", cConnect)
'    FrmAddOCOtrosCliente.Show vbModal
'    Set FrmAddOCOtrosCliente = Nothing
'    Call CARGA_GRID
''Case "ACTUALIZAR"
''    For i = 1 To GridEX1.RowCount
''        GridEX1.Row = i
''        Call Actualizar_Precios
''    Next
''    Call CARGA_GRID
'Case "SALIR"
'    Unload Me
'End Select
'End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
If ColIndex <> GridEX1.Columns("Precio").Index Then
    Cancel = True
End If
End Sub

Sub Actualizar_Precios()
On Error GoTo errActualizar

strSQL = "Ventas_Up_Actualiza_Precio_Tejeduria_Otros_Clientes '" & sCod_Cliente & "','" & sSer_OrdComp & "','" & sCod_OrdComp & "','" & sSec_OrdComp & "','" & GridEX1.Value(GridEX1.Columns("cod_otro_cliente").Index) & "'," & GridEX1.Value(GridEX1.Columns("Precio").Index)
ExecuteCommandSQL cConnect, strSQL

Exit Sub
errActualizar:
    MsgBox Err.Description, vbCritical, "Actualizar Precios"
End Sub

Private Sub GridEX1_GotFocus()
'GridEX1.Col = GridEX1.Columns("Precio").Index
End Sub

Private Sub GridEX1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    GridEX1.Row = GridEX1.Row + 1
End If
End Sub
