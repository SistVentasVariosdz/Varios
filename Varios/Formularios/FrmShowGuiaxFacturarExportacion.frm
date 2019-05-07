VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowGuiaxFacturarExportacion 
   Caption         =   "Facturación Guias Cliente Exportación"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   16050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_modificar 
      Caption         =   "Modificar"
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txt_maxguiapermitido 
      Height          =   285
      Left            =   2760
      TabIndex        =   33
      Top             =   7440
      Width           =   1095
   End
   Begin GridEX20.GridEX GridEX4 
      Height          =   2055
      Left            =   6390
      TabIndex        =   29
      Top             =   4920
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmShowGuiaxFacturarExportacion.frx":0000
      Column(2)       =   "FrmShowGuiaxFacturarExportacion.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmShowGuiaxFacturarExportacion.frx":016C
      FormatStyle(2)  =   "FrmShowGuiaxFacturarExportacion.frx":02A4
      FormatStyle(3)  =   "FrmShowGuiaxFacturarExportacion.frx":0354
      FormatStyle(4)  =   "FrmShowGuiaxFacturarExportacion.frx":0408
      FormatStyle(5)  =   "FrmShowGuiaxFacturarExportacion.frx":04E0
      FormatStyle(6)  =   "FrmShowGuiaxFacturarExportacion.frx":0598
      ImageCount      =   0
      PrinterProperties=   "FrmShowGuiaxFacturarExportacion.frx":0678
   End
   Begin GridEX20.GridEX GridEX3 
      Height          =   2055
      Left            =   3360
      TabIndex        =   30
      Top             =   4950
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmShowGuiaxFacturarExportacion.frx":0850
      Column(2)       =   "FrmShowGuiaxFacturarExportacion.frx":0918
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmShowGuiaxFacturarExportacion.frx":09BC
      FormatStyle(2)  =   "FrmShowGuiaxFacturarExportacion.frx":0AF4
      FormatStyle(3)  =   "FrmShowGuiaxFacturarExportacion.frx":0BA4
      FormatStyle(4)  =   "FrmShowGuiaxFacturarExportacion.frx":0C58
      FormatStyle(5)  =   "FrmShowGuiaxFacturarExportacion.frx":0D30
      FormatStyle(6)  =   "FrmShowGuiaxFacturarExportacion.frx":0DE8
      ImageCount      =   0
      PrinterProperties=   "FrmShowGuiaxFacturarExportacion.frx":0EC8
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   2055
      Left            =   360
      TabIndex        =   31
      Top             =   4950
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmShowGuiaxFacturarExportacion.frx":10A0
      Column(2)       =   "FrmShowGuiaxFacturarExportacion.frx":1168
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmShowGuiaxFacturarExportacion.frx":120C
      FormatStyle(2)  =   "FrmShowGuiaxFacturarExportacion.frx":1344
      FormatStyle(3)  =   "FrmShowGuiaxFacturarExportacion.frx":13F4
      FormatStyle(4)  =   "FrmShowGuiaxFacturarExportacion.frx":14A8
      FormatStyle(5)  =   "FrmShowGuiaxFacturarExportacion.frx":1580
      FormatStyle(6)  =   "FrmShowGuiaxFacturarExportacion.frx":1638
      ImageCount      =   0
      PrinterProperties=   "FrmShowGuiaxFacturarExportacion.frx":1718
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   615
      Left            =   14400
      TabIndex        =   28
      Top             =   240
      Width           =   1425
      _ExtentX        =   2540
      _ExtentY        =   1111
      Custom          =   "0~0~Salir~Verdadero~Verdadero~&Salir~1~0~~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1400
      ControlHeigth   =   600
      ControlSeparator=   50
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6120
      Left            =   0
      TabIndex        =   12
      Top             =   1170
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   10795
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
      FormatStyle(1)  =   "FrmShowGuiaxFacturarExportacion.frx":18F0
      FormatStyle(2)  =   "FrmShowGuiaxFacturarExportacion.frx":1A28
      FormatStyle(3)  =   "FrmShowGuiaxFacturarExportacion.frx":1AD8
      FormatStyle(4)  =   "FrmShowGuiaxFacturarExportacion.frx":1B8C
      FormatStyle(5)  =   "FrmShowGuiaxFacturarExportacion.frx":1C64
      FormatStyle(6)  =   "FrmShowGuiaxFacturarExportacion.frx":1D1C
      FormatStyle(7)  =   "FrmShowGuiaxFacturarExportacion.frx":1DFC
      ImageCount      =   0
      PrinterProperties=   "FrmShowGuiaxFacturarExportacion.frx":1E1C
   End
   Begin VB.Frame fraPrecio 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Modificación"
      Height          =   2520
      Left            =   10500
      TabIndex        =   18
      Top             =   1890
      Visible         =   0   'False
      Width           =   3030
      Begin VB.TextBox txtImp_comision 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1590
         TabIndex        =   25
         Text            =   "0"
         Top             =   1260
         Width           =   1125
      End
      Begin VB.TextBox txtPorc_Descuento_Precio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         TabIndex        =   20
         Text            =   "0"
         Top             =   390
         Width           =   540
      End
      Begin VB.TextBox txtPre_Unitario 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         TabIndex        =   21
         Text            =   "0"
         Top             =   825
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancelarPrecio 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   1545
         TabIndex        =   24
         Top             =   1770
         Width           =   990
      End
      Begin VB.CommandButton cmdAceptarPrecio 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   495
         TabIndex        =   22
         Top             =   1770
         Width           =   990
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Importe Comisión"
         Height          =   300
         Left            =   135
         TabIndex        =   26
         Top             =   1305
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "% Descuento Precio"
         Height          =   435
         Left            =   135
         TabIndex        =   23
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Precio Unitario"
         Height          =   315
         Left            =   135
         TabIndex        =   19
         Top             =   870
         Width           =   1485
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15915
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   9480
         TabIndex        =   27
         Tag             =   "..."
         Top             =   255
         Width           =   300
      End
      Begin VB.TextBox txtCod_TipoFact 
         Height          =   315
         Left            =   7560
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   9765
         TabIndex        =   1
         Top             =   240
         Width           =   3075
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   8760
         TabIndex        =   0
         Top             =   240
         Width           =   570
      End
      Begin VB.CheckBox optTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Cbo_Almacen 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   4080
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90570753
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   3450
         TabIndex        =   9
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90570753
         CurrentDate     =   37543
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   570
         Left            =   12960
         TabIndex        =   4
         Top             =   240
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1005
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~~~0~Verdadero~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1100
         ControlHeigth   =   550
         ControlSeparator=   50
      End
      Begin VB.TextBox txtDes_TipoFact 
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo de Facturación"
         Height          =   420
         Left            =   6480
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   7800
         TabIndex        =   16
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Rango Fecha de Emisión:"
         Height          =   360
         Left            =   105
         TabIndex        =   11
         Top             =   645
         Width           =   1710
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Máximo Nro de Guias Permitidas"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   7440
      Width           =   2415
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6255
      Top             =   2955
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label lbRollos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8730
      TabIndex        =   15
      Top             =   6690
      Width           =   45
   End
   Begin VB.Label lbDes_Color 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9690
      TabIndex        =   14
      Top             =   6690
      Width           =   45
   End
   Begin VB.Label lbGuia 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9690
      TabIndex        =   13
      Top             =   6990
      Width           =   45
   End
End
Attribute VB_Name = "FrmShowGuiaxFacturarExportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim Doc As String
Dim strSQL As String
Public codigo As String
Public Descripcion As String
Public TipoAdd As String
Dim sCod_TipoFact  As String

Dim sSer_Factura_Orig As String
Dim sNum_Factura_Orig As String

Private Sub DtFecVencimiento_Change()
  GridEX1.ClearFields
  dtpFecEmiIni.Value = ""
  dtpFecEmiFin.Value = ""
End Sub

Private Sub cmd_modificar_Click()

If IsNumeric(txt_maxguiapermitido.Text) = True Then
    If txt_maxguiapermitido.Text > 0 Then
        If MsgBox("Estado seguro de cambiar el máximo de guias a ingresar?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
            Call Modifica_NroRegistro
        End If
    Else
    MsgBox "Debe Ingresar una cantidad diferente de cero", vbInformation, "Información"
    End If
Else
MsgBox "Debe Ingresar una cantidad", vbInformation, "Información"
End If


End Sub
Sub Modifica_NroRegistro()
Dim Rs As ADODB.Recordset

Dim i As Integer
On Error GoTo hand

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient

Rs.Open "Exec Usp_Upd_Actualiza_Parametro_001 " & txt_maxguiapermitido.Text & ",'" & vusu & "'"

MsgBox "Se actualizo el máximo de guias por factura", vbInformation, Me.Caption

Exit Sub
hand:
    ErrorHandler err, "ACTUALIZA NRO DE GUIA POR FACTURA"
    Set Rs = Nothing
End Sub
Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oPARENT = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TX_Cliente  ORDER BY Abr_Cliente"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtAbr_Cliente.Text = codigo
        txtNom_Cliente.Text = Descripcion
        strSQL = "SELECT Cod_Cliente_Tex As Cod_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        txtAbr_Cliente.Tag = DevuelveCampo(strSQL, cConnect)

        SendKeys "{TAB}"
        codigo = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
End Sub

Private Sub Form_Load()

  dtpFecEmiIni.Value = Date
  dtpFecEmiFin.Value = Date
  
  FillAlmacen
  
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If

End Sub

Private Sub Buscar()

On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

   
sSQL = "EXEC TI_DESPACHOS_TELATINTORERIA_MARCA_MULTIP_FAC_EXPORTACION '" & Left(Cbo_Almacen, 2) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & Trim(txtAbr_Cliente.Tag) & "'"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
  
Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)


GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = False
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

GridEX1.Columns("Fecha_Guia").Width = 900
GridEX1.Columns("Fecha_Guia").EditType = jgexEditNone

GridEX1.Columns("Ser_Factura").Width = 400
GridEX1.Columns("Num_Factura").Width = 1100
GridEX1.Columns("cliente").Width = 2000
GridEX1.Columns("cliente").EditType = jgexEditNone

GridEX1.Columns("Guia").Width = 1260
GridEX1.Columns("Guia").EditType = jgexEditNone

GridEX1.Columns("DES_TIPMOV").Width = 1800
GridEX1.Columns("DES_TIPMOV").EditType = jgexEditNone

GridEX1.Columns("TipoServicio").Width = 1800
GridEX1.Columns("TipoServicio").EditType = jgexEditNone

GridEX1.Columns("COD_USUARIO").Width = 1000
GridEX1.Columns("COD_USUARIO").Visible = False

'
GridEX1.Columns("KILOS_TOTALES").Visible = True
GridEX1.Columns("KILOS_TOTALES").Width = 700
GridEX1.Columns("KILOS_TOTALES").EditType = jgexEditNone

GridEX1.Columns("ROLLOS_TOTALES").Width = 700
GridEX1.Columns("ROLLOS_TOTALES").EditType = jgexEditNone

GridEX1.Columns("FLG_STATUS").Width = 680
GridEX1.Columns("FLG_STATUS").Visible = False

GridEX1.Columns("OBSERVACIONES").Width = 3000
GridEX1.Columns("OBSERVACIONES").EditType = jgexEditNone

GridEX1.Columns("NUM_MOVSTK").Width = 1500
GridEX1.Columns("NUM_MOVSTK").EditType = jgexEditNone

GridEX1.Columns("FAC_CLI").Width = 580
GridEX1.Columns("FAC_CLI").Visible = False



GridEX1.Columns("COD_CLIENTE_TEX").Width = 580
GridEX1.Columns("COD_CLIENTE_TEX").Visible = False

SetColores

GridEX1.DefaultGroupMode = jgexDGMCollapsed

If dtpFecEmiIni.Value <> "" Then
    GridEX1.DefaultGroupMode = jgexDGMExpanded
End If

If GridEX1.RowCount > 0 Then
    GridEX1.Row = 1
End If

GridEX1.ContinuousScroll = True

Exit Sub
Resume
drDepurar:
  Errores err.Number
End Sub

Private Sub Form_Resize()
'GridEX1.Width = Me.Width - 300
GridEX1.Height = Me.Height - 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      Buscar
    End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "SALIR"
       Unload Me
    End Select
End Sub


Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
  'If Left(Cbo_Almacen, 2) = "TT" Then
    If ColIndex = 3 Then '--NUMERO DE FACTURA
      AfterColEdit_Prendas (ColIndex)
    End If
  'End If
End Sub
'kkk
Sub AfterColEdit_Prendas(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup


Select Case ColIndex
   
    
  Case Is = GridEX1.Columns("Num_Factura").Index
    Call Asigna_Numero_Factura
    Buscar
    
   
  End Select
  
  
Exit Sub

Resume

Error_Handler:

  Errores err.Number
   
  If ColIndex = GridEX1.Columns("Sel").Index Then
     GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End If
End Sub
Private Sub Asigna_Numero_Factura()
Dim sSQL As String
Dim num_factura As String
Dim Serie As String
On Error GoTo errx
    Serie = "000"
    Serie = Serie + Replace(Trim(GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)), " ", "")
    Serie = Right(Serie, 3)

    num_factura = "00000000"
    num_factura = num_factura + Replace(Trim(GridEX1.Value(GridEX1.Columns("Num_Factura").Index)), " ", "")
    num_factura = Right(num_factura, 8)
   

      sSQL = "USP_INS_ACT_FACTURA_VENTA_EXPORTACION '$','$','$','$','$'"
      
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       Serie, _
                       num_factura, _
                       GridEX1.Value(GridEX1.Columns("Cod_Cliente_Tex").Index))


    ExecuteCommandSQL cConnect, sSQL
    
    
    
Exit Sub
errx:
    Errores err.Number
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

If Left(Cbo_Almacen, 2) = "TT" Then
  Select Case ColIndex
'    Case Is = GridEX1.Columns("Ser_Factura").Index
'        sSer_Factura_Orig = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'        sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
'        'sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
'        Cancel = False
    Case Is = GridEX1.Columns("Num_Factura").Index
        sSer_Factura_Orig = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
        sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
        Cancel = False
    Case Is = GridEX1.Columns("SEL").Index
      Cancel = False
      
   Case Else
      Cancel = True
    End Select
End If
  
End Sub

Private Sub GridEX1_Click()

'On Error Resume Next
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String
    
    bCargaGRid = True
    
        If GridEX1.RowCount > 0 Then
        ColIndex = GridEX1.Col
        
        If Not GridEX1.IsGroupItem(GridEX1.Row) And ColIndex > 0 Then
            If UCase(GridEX1.Columns(ColIndex).Key) = "SEL" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
        Else
            If GridEX1.IsGroupItem(GridEX1.Row) Then
            End If
        End If
    End If
End Sub

Private Sub GridEX1_DblClick()
'    Dim i As Integer
'    For i = 1 To GridEX1.Columns.Count
'        Debug.Print GridEX1.Name & ".Columns(" & Chr(34) & GridEX1.Columns(i).Key & Chr(34) & ").width = " & CStr(GridEX1.Columns(i).Width)
'    Next
'
'    For i = 1 To GridEX1.Columns.Count
'        Debug.Print GridEX1.Name & ".COLUMNS(" & Chr(34) & GridEX1.Columns(i).Key & Chr(34) & ").CAPTION = " & CStr(GridEX1.Columns(i).Caption)
'    Next
    
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Dim ocol As JSColumn
    Dim oRow As JSRowData
    Dim vCurrentRow As Variant
    Dim oRowGroup As JSRowData
    Dim sProveedor As String
    
    iColAnterior = LastCol
    iRowAnterior = LastRow
    
    If GridEX1.Row <> 0 Then
        Set oRow = GridEX1.GetRowData(GridEX1.Row)
    End If
      
    If GridEX1.RowCount > 0 Then
      On Error Resume Next
    End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)

Dim strGroupCaption As String

If RowBuffer.RowType = jgexRowTypeGroupHeader Then
    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
    RowBuffer.GroupCaption = strGroupCaption
End If

End Sub

'Private Sub MuestraSubTotales()
'Dim colTemp As JSColumn
'
'GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
'Set colTemp = GridEX1.Columns("Moneda")
'colTemp.AggregateFunction = jgexAggregateNone
'colTemp.TotalRowPrefix = "SUB TOTAL "
'
'GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
'Set colTemp = GridEX1.Columns("Num_Prendas")
'colTemp.AggregateFunction = jgexSum
'colTemp.TotalRowPrefix = ""
'
'GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
'Set colTemp = GridEX1.Columns("MontoDespacho")
'colTemp.AggregateFunction = jgexSum
'colTemp.TotalRowPrefix = ""
'
'End Sub

Private Sub SetColores()

Dim fmtCon As JSFmtCondition
Dim fmtCond2 As JSFmtCondition
Dim fmtCond3 As JSFmtCondition

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("SEL").Index, jgexEqual, -1)
    
    With GridEX1.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
            Set fmtCon = .GroupCondition
    End With
    fmtCon.SetCondition GridEX1.Columns("SEL").Index, jgexEqual, -1
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
End Sub


'Private Sub Autorizar()
'
'On Error GoTo errorx
'Dim sSQL As String
'Dim aMess(4), i As Integer
'
'
'GridEX1.MoveFirst
'
'For i = 0 To GridEX1.RowCount
'
'  If GridEX1.Value(GridEX1.Columns("SEL").Index) Then
'
'    If Left(Cbo_Almacen, 2) = "62" Then
'
'      'ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$ ,'$','$','$','$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$',$,$"
'      sSQL = "Ventas_Cambio_Estado_DocAlm_Prendas_Clientes_Locales'$','$','$','$','$',$,'$',$,$ ,'$','$','$','$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$',$,$"
'
'      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
'                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
'                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
'                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
'                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
'                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
'                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
'                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
'                       GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), GridEX1.Value(GridEX1.Columns("Imp_DESCUENTO").Index), GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
'                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
'                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_Comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))
'
'
'
'      ExecuteCommandSQL cConnect, sSQL
'    End If
'  End If
'
'  GridEX1.MoveNext
'
'Next i
'
'If Left(Cbo_Almacen, 2) = "62" Then
'  'ExecuteCommandSQL cCONNECT, "Ventas_Genera_Docum_Autorizados_Prendas '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
'
'  ExecuteCommandSQL cConnect, "Ventas_Genera_Docum_Autorizados_Prendas_Clientes_locales '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
'
'End If
'
'Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
'
'BUSCAR
'
'Exit Sub
'Resume
'errorx:
'    Errores err.Number
'End Sub

'Sub Cambio_Nro_Factura()
'
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  lvSw = True
'
'  Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index)
'  serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index)
'
'  GridEX1.MoveFirst
'  For i = 0 To GridEX1.RowCount
'    If Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'      GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) = serie
'      GridEX1.Value(GridEX1.Columns("Nro_Docum_Ventas").Index) = Nro_Factura
'    End If
'    GridEX1.MoveNext
'  Next i
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub


'
'Private Sub GridEX2_Click()
'
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  lvSw = True
'
'  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'
'
'  GridEX1.MoveFirst
'  For i = 0 To GridEX1.RowCount
'    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'      GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index)
'      GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = GridEX2.Value(GridEX2.Columns("Descripcion").Index)
'    End If
'    GridEX1.MoveNext
'  Next i
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub

'Private Sub GridEX3_Click()
'
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'  lvSw = True
'  GridEX1.MoveFirst
'  For i = 0 To GridEX1.RowCount
'    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'      GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index) = GridEX3.Value(GridEX3.Columns("Cod_Moneda").Index)
'      GridEX1.Value(GridEX1.Columns("Moneda").Index) = GridEX3.Value(GridEX3.Columns("Descripcion").Index)
'    End If
'    GridEX1.MoveNext
'  Next i
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub


Private Sub FillAlmacen()

Dim rstAux As ADODB.Recordset
Dim strSQL As String
    

strSQL = "EXEC LG_MUESTRA_ALMACENES_TX_POR_USUARIO_FAC_EXPORTACION '" & vusu & "'"
    
         
Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
Cbo_Almacen.Clear
With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        Cbo_Almacen.AddItem !COD_ALMACEN & " " & !Nom_Almacen
        .MoveNext
    Loop
    .Close
End With
If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
Set rstAux = Nothing
    
End Sub



'Private Sub GridEX4_Click()
'
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  lvSw = True
'
'  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'
'
'  GridEX1.MoveFirst
'  For i = 0 To GridEX1.RowCount
'    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'      GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) = GridEX4.Value(GridEX4.Columns("Cod_Anxo").Index)
'      GridEX1.Value(GridEX1.Columns("Des_Anexo").Index) = GridEX4.Value(GridEX4.Columns("Des_Anexo").Index)
'
'      If RTrim(FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), vbString)) = "" Then
'        GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = GridEX4.Value(GridEX4.Columns("Pie_Factura1").Index)
'      End If
'      If RTrim(FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), vbString)) = "" Then
'        GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = GridEX4.Value(GridEX4.Columns("Pie_Factura2").Index)
'      End If
'
'    End If
'    GridEX1.MoveNext
'  Next i
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'
'  SendKeys "{TAB}"
'
'End Sub

'Public Sub BuscaCliente(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSql = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
'
'    txtAbr_Cliente = Trim(txtAbr_Cliente)
'    txtNom_Cliente = Trim(txtNom_Cliente)
'
'    Select Case Opcion
'    Case 1: strSql = strSql & "Abr_Cliente LIKE '%" & txtAbr_Cliente & "%'"
'    Case 2: strSql = strSql & "Nom_Cliente LIKE '%" & txtNom_Cliente & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSql
'    frmBusqGeneral3.Cargar_Datos
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Cliente").Visible = False
'    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Width = 570
'    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Width = 2370
'
'    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Caption = "Abrev."
'    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Caption = "Cliente"
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtAbr_Cliente.Tag = ""
'    txtAbr_Cliente = ""
'    txtNom_Cliente = ""
'    If Codigo <> "" Then
'
'        txtAbr_Cliente = Descripcion
'        txtNom_Cliente = TipoAdd
'        txtAbr_Cliente.Tag = Codigo
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    Codigo = ""
'    Descripcion = ""
'End Sub




'Private Sub txtCod_Class_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtCod_Embarque_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaModoTransporte 1
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtCod_Vendor_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

'Private Sub txtImp_comision_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        cmdAceptarPrecio.SetFocus
'    End If
'End Sub
'
'Private Sub txtNom_embarque_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
'End Sub

Private Sub txtAbr_Cliente_Change()
        txtAbr_Cliente.Tag = ""
    
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaCliente 1
'        SendKeys "{TAB}"
'    End If
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cConnect)
            strSQL = "SELECT Cod_Cliente_Tex As Cod_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtAbr_Cliente.Tag = DevuelveCampo(strSQL, cConnect)
            
            

            SendKeys "{TAB}"


        End If
    End If
End Sub

'Private Sub txtCartaCredito_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        BuscaCartaCredito 1
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtImp_Descuento_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'       SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtImp_Flete_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
'
'End Sub
'
'Private Sub txtImp_Seguro_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
'End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaCliente 2
'        SendKeys "{TAB}"
'    End If

    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TX_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cConnect)
            strSQL = "SELECT Nom_Cliente FROM TX_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cConnect)
            strSQL = "SELECT Cod_Cliente_Tex FROM TX_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtAbr_Cliente.Tag = DevuelveCampo(strSQL, cConnect)
            SendKeys "{TAB}"

        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtNom_Cliente.SetFocus
        End If
    End If
End Sub

Private Sub CargarDatos()

'    txtobservacion.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), vbString)
'    txtSecuencia.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), vbLong)
'    txtLinea1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index), vbString)
'    txtCod_CondVent.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), vbString)
'    txtDes_CondVent.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index), vbString)
'    txtCartaCredito.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString)
'    txtImp_Flete.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), vbDouble)
'    txtImp_Seguro.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), vbDouble)
'    txtImp_Descuento.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), vbDouble)
'    txtCod_Termino_Venta = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), vbString)
'    txtDes_Termino_Venta = FixNulos(GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index), vbString)
'    txtCod_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Embarque").Index), vbString)
'    txtDes_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Des_Embarque").Index), vbString)
'    txtNom_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), vbString)
'    txtPie_Pagina1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), vbString)
'    txtPie_Pagina2.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), vbString)
'    txtCod_Vendor.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), vbString)
'    txtCod_Class.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Class").Index), vbString)
'    txtPor_Comision.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Por_Comision").Index), vbDouble)
'
'    txtRef_Embarque.Text = FixNulos(DevuelveCampo("select ref_embarque FROM TG_EMBARQUE where num_embarque = '" & FixNulos(GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vbLong) & "'", cCONNECT), vbString)
'
'    txtImp_Desaduanaje.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index), vbDouble)
'    txtImp_Transporte_Pais_Destino.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index), vbDouble)
'
'
    'Me.fraDatosAdicionales.Visible = True
    'Me.txtRef_Embarque.SetFocus
End Sub

'Private Sub GuardarDatos()
'On Error GoTo errx
'Dim sSQL As String

'    GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtobservacion.Text
'    GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
'    GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1
'    GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
'    GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
'    GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete
'    GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
'    GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
'    GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
'    GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
'    GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
'    GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
'    GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
'    GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
'    GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
'    GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cCONNECT), vbLong)
'    GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text

'      ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$,'$','$','$' ,'$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$',$,$"
'
'      ssql = VBsprintf(ssql, Left(Cbo_Almacen, 2), _
'                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
'                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
'                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
'                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
'                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
'                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
'                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
'                       GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), _
'                       GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
'                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
'                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))
'
'
'    ExecuteCommandSQL cCONNECT, ssql
'
'    DatosAdic_Click
'
'    GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtobservacion.Text
'    GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
'    GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1
'    GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
'    GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
'    GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete
'    GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
'    GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
'    GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
'    GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
'    GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
'    GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
'    GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
'    GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
'    GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
'    GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cCONNECT), vbLong)
'    GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
'    GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
'
    'Me.fraDatosAdicionales.Visible = False
'Exit Sub
'errx:
'    Errores err.Number
'End Sub

'Private Sub DatosAdic_Click()
'
'Dim Serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  lvSw = True
'
'  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'
'
'  GridEX1.MoveFirst
'  For I = 0 To GridEX1.RowCount
'    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'        GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtobservacion.Text
'        GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
'        GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1.Text
'        GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
'        GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
'        GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
'        GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
'        GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
'        GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
'        GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
'        GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
'        GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
'        GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
'        GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
'        GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cCONNECT), vbLong)
'        GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
'    End If
'    GridEX1.MoveNext
'  Next I
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'
'End Sub


'Private Sub txtobservacion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtPie_Pagina1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtPie_Pagina2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
'End Sub

'Private Sub txtPorc_Descuento_Precio_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn And txtPorc_Descuento_Precio > 0 Then
'        txtPre_Unitario.Text = GridEX1.Value(GridEX1.Columns("Pre_Unitario_ORG").Index) - Round(GridEX1.Value(GridEX1.Columns("Pre_Unitario_ORG").Index) * (Val(txtPorc_Descuento_Precio) / 100), 2)
'        cmdAceptarPrecio.SetFocus
'    End If
'End Sub

'Private Sub txtPre_Unitario_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtImp_comision.SetFocus
'    End If
'End Sub
'
'Private Sub txtRef_Embarque_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaRef_Embarque 1
'        SendKeys "{TAB}"
'    End If
'
'End Sub
'
'Private Sub txtSecuencia_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'
'        BuscaLugEnt 1
'        SendKeys "{TAB}"
'    End If
'End Sub

'Public Sub BuscaLugEnt(Opcion As String)
'Dim rstAux As ADODB.Recordset
'    strSQL = "SELECT Secuencia, RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
'             "RTRIM(Linea3) AS Linea1 FROM TG_CLIENTE_LUGENT " & _
'             "WHERE Cod_Cliente = '" & txtAbr_Cliente.Tag & "' AND "
'
'    txtSecuencia = Trim(txtSecuencia)
'    txtLinea1 = Trim(txtLinea1)
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "CONVERT(varchar(8), Secuencia) like '%" & txtSecuencia & "%'"
'    Case 2: strSQL = strSQL & "RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
'             "RTRIM(Linea3) LIKE '%" & txtLinea1 & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Secuencia").Visible = False
'    frmBusqGeneral3.gexLista.Columns("Secuencia").Width = 570
'    frmBusqGeneral3.gexLista.Columns("Linea1").Width = 2370
'
'    frmBusqGeneral3.gexLista.Columns("Secuencia").Caption = "Secuencia"
'    frmBusqGeneral3.gexLista.Columns("Linea1").Caption = "Lug.Entr."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtSecuencia = ""
'    txtLinea1 = ""
'
'    If Codigo <> "" Then
'        txtSecuencia = Codigo
'        txtLinea1 = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    Codigo = ""
'    Descripcion = ""
'End Sub




'Private Sub txtCod_CondVent_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaCondVent 1
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Public Sub BuscaCondVent(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSQL = "SELECT Cod_CondVent, Des_CondVent FROM lg_condvent WHERE "
'
'    txtCod_CondVent = Trim(txtCod_CondVent)
'    txtDes_CondVent = Trim(txtDes_CondVent)
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "Cod_condVent like '%" & txtCod_CondVent & "%'"
'    Case 2: strSQL = strSQL & "Des_condVent LIKE '%" & txtDes_CondVent & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'
'    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Width = 700
'    frmBusqGeneral3.gexLista.Columns("Des_CondVent").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Caption = "Cond.Vta"
'    frmBusqGeneral3.gexLista.Columns("Des_condVent").Caption = "Descrip."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_CondVent = ""
'    txtDes_CondVent = ""
'
'    If Codigo <> "" Then
'        txtCod_CondVent = Codigo
'        txtDes_CondVent = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    Codigo = ""
'    Descripcion = ""
'End Sub
'
'
'Private Function Busca_AnexosCliente()
'  If GridEX1.RowCount > 0 Then
'      Set GridEX4.ADORecordset = CargarRecordSetDesconectado("SM_TG_CLIENTE_ANEXOCONT '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "'", cConnect)
'      GridEX4.ColumnAutoResize = True
'
'
'      GridEX4.Columns("COD_CLIENTE").Visible = False
'      GridEX4.Columns("COD_ANXO").Visible = False
'      GridEX4.Columns("COD_TIPANEX").Visible = False
'
'        GridEX4.ActAsDropDown = True
'        GridEX4.BoundColumnIndex = 2
'        GridEX4.ReplaceColumnIndex = 2
'
'  End If
'
'End Function

'Private Sub txtCod_TipoFact_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaTipoFacturacion 1
'        FunctButt1.SetFocus
'    End If
'End Sub

'Public Sub BuscaTipoFacturacion(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSql = "SELECT Cod_TipoFact, Des_TipoFact FROM CN_TipoFactura_Venta WHERE "
'
'    txtCod_TipoFact = Trim(txtCod_TipoFact)
'    txtDes_TipoFact = Trim(txtDes_TipoFact)
'
'    Select Case Opcion
'    Case 1: strSql = strSql & "Cod_TipoFact LIKE '%" & txtCod_TipoFact & "%'"
'    Case 2: strSql = strSql & "Des_TipoFact LIKE '%" & txtDes_TipoFact & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSql
'    frmBusqGeneral3.Cargar_Datos
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Cod_tipoFact").Width = 800
'    frmBusqGeneral3.gexLista.Columns("Des_TipoFact").Width = 10000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_tipoFact").Caption = "Tipo de Facturación"
'    frmBusqGeneral3.gexLista.Columns("des_tipoFact").Caption = "Descripción de Facturación "
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_TipoFact.Tag = ""
'    txtCod_TipoFact = ""
'    txtDes_TipoFact = ""
'
'    If Codigo <> "" Then
'
'        txtCod_TipoFact = Codigo
'        txtDes_TipoFact = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    Codigo = ""
'    Descripcion = ""
'End Sub


'Private Sub SeleccionarOtrosReg(valor As Variant)
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'Dim sSQL As String
'  GridEX1.Redraw = False
'
'  lvSw = True
'
'  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'
'
'  GridEX1.MoveFirst
'  For i = 0 To GridEX1.RowCount
'    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'        GridEX1.Value(GridEX1.Columns("Sel").Index) = valor
'      'ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$,'$','$','$' ,'$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$'"
'      sSQL = "Ventas_Cambio_Estado_DocAlm_Prendas_Clientes_Locales '$','$','$','$','$',$,'$',$,$,'$','$','$' ,'$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$'"
'
'      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
'                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
'                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
'                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
'                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
'                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
'                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
'                       GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
'                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
'                       GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
'                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), GridEX1.Value(GridEX1.Columns("Imp_DESCUENTO").Index), GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
'                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
'                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
'                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_Comision").Index))
'      ExecuteCommandSQL cConnect, sSQL
'
'    End If
'    GridEX1.MoveNext
'  Next i
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'End Sub


'Private Sub txtCod_Termino_Venta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaTerminoVent 1
'        SendKeys "{TAB}"
'    End If
'End Sub

'Public Sub BuscaTerminoVent(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSQL = "SELECT Cod_Termino_Venta, Des_Termino_Venta FROM CN_Termino_Venta WHERE "
'
'    txtCod_Termino_Venta = Trim(txtCod_Termino_Venta)
'    txtDes_Termino_Venta = Trim(txtDes_Termino_Venta)
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "Cod_Termino_Venta like '%" & txtCod_Termino_Venta & "%'"
'    Case 2: strSQL = strSQL & "Des_Termino_Venta LIKE '%" & txtDes_Termino_Venta & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Width = 700
'    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Caption = "Termino.Venta"
'    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Caption = "Descrip."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_Termino_Venta = ""
'    txtDes_Termino_Venta = ""
'
'    If Codigo <> "" Then
'        txtCod_Termino_Venta = Codigo
'        txtDes_Termino_Venta = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    Codigo = ""
'    Descripcion = ""
'End Sub







Public Function CargaValores(ByRef ObjTemp As Object) As Boolean
    ObjTemp.txtAbr_Cliente.Text = txtAbr_Cliente.Text
    ObjTemp.txtAbr_Cliente.Tag = txtAbr_Cliente.Tag
    ObjTemp.txtDes_Cliente.Text = txtNom_Cliente.Text
    'ObjTemp.txtCOD_TEMCLI.Text = gexLista.Value(gexLista.Columns("COD_TEMCLI").Index)
    'ObjTemp.CARGA_ESTCLI
End Function


Private Sub Cambio_Fecha(SFecha As String)
Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
Dim sSQL As String
  GridEX1.Redraw = False

  lvSW = True
  
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
        GridEX1.Value(GridEX1.Columns("Fecha").Index) = SFecha
    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True

End Sub



'Private Sub Cambio_PO_Factura(sPO As String)
'Dim sSQL As String
'On Error GoTo errx
'
'    GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index) = sPO
'
'    sSQL = "UP_MAN_TEMP_Ventas_PurOrd_Factura_Clientes_locales '$','$','$','$','$',$,'$','$','$','$','$','$','$'"
'
'    sSQL = VBsprintf(sSQL, "I", vusu, Left(Cbo_Almacen, 2), _
'            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
'            GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
'            GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_PurOrd").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_LotPurOrd").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_Estcli").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_ColCli").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_Talla").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index))
'
'    ExecuteCommandSQL cConnect, sSQL
'
'Exit Sub
'errx:
'    Errores err.Number
'
'End Sub



'Private Function GrabaDatosParaFacturaCambiada() As Boolean
'Dim sSQL As String
'Dim num_factura As String
'Dim serie As String
'On Error GoTo errx
'    serie = "000"
'    serie = serie + Replace(Trim(GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)), " ", "")
'    serie = Right(serie, 3)
'
'    num_factura = "00000000"
'    num_factura = num_factura + Replace(Trim(GridEX1.Value(GridEX1.Columns("Num_Factura").Index)), " ", "")
'    num_factura = Right(num_factura, 8)
'
'
'      sSQL = "USP_INS_ACT_FACTURA_VENTA '$','$','$','$','$'"
'
'      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
'                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
'                       serie, _
'                       num_factura, _
'                       GridEX1.Value(GridEX1.Columns("Cod_Cliente_Tex").Index))
'
'
'    ExecuteCommandSQL cConnect, sSQL
'
'    GrabaDatosParaFacturaCambiada = True
'
'Exit Function
'errx:
'    Errores err.Number
'    GrabaDatosParaFacturaCambiada = False
'End Function



'
'Public Sub BuscaCartaCredito(Opcion As String)
'Dim rstAux As ADODB.Recordset
'    strSQL = "SELECT Num_CartaCredito , Fec_Emision " & _
'             "FROM TG_Carta_Credito " & _
'             "WHERE Cod_Cliente = '" & txtAbr_Cliente.Tag & "' AND "
'
'    txtCartaCredito = Trim(txtCartaCredito)
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "Num_CartaCredito like '%" & txtCartaCredito & "%'"
'    End Select
'    strSQL = strSQL & " AND FLG_STATUS IN ('B','F','T')"
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Visible = True
'    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Width = 2000
'    frmBusqGeneral3.gexLista.Columns("Fec_Emision").Width = 1500
'
'    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Caption = "Carta Credito"
'    frmBusqGeneral3.gexLista.Columns("Fec_Emision").Caption = "Fec_Emision"
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    If Codigo <> "" Then
'        txtCartaCredito = Codigo
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    Codigo = ""
'    Descripcion = ""
'End Sub


