VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmShowFactVentas 
   Caption         =   "Documentos Ventas"
   ClientHeight    =   8970
   ClientLeft      =   180
   ClientTop       =   495
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   14385
   Begin VB.Frame FraBuscar 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12720
      Begin VB.OptionButton optPartida 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Partida"
         Height          =   375
         Left            =   5280
         TabIndex        =   41
         Top             =   120
         Width           =   1695
      End
      Begin VB.Frame frPartida 
         BackColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   9285
         Begin VB.TextBox txtPartidaBusqueda 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   960
            MaxLength       =   11
            TabIndex        =   39
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Partida:"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.OptionButton optNroDoc 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nro de Documento"
         Height          =   375
         Left            =   3120
         TabIndex        =   37
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton optAnoMes 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Año/ Mes"
         Height          =   375
         Left            =   1560
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton optCliente 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Anexo"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   1215
      End
      Begin VB.Frame frCliente 
         BackColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   9285
         Begin VB.TextBox txtCod_TipAnxo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6960
            MaxLength       =   1
            TabIndex        =   32
            Text            =   "C"
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtDes_Anexo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   31
            Top             =   240
            Width           =   4050
         End
         Begin VB.TextBox txtNum_Ruc 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   960
            MaxLength       =   11
            TabIndex        =   30
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   6480
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nro Ruc:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame frAnoMes 
         BackColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   9285
         Begin VB.TextBox txtAno 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   26
            Top             =   345
            Width           =   660
         End
         Begin VB.TextBox txtMes 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   25
            Top             =   345
            Width           =   480
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Mes"
            Height          =   255
            Left            =   3960
            TabIndex        =   28
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Año"
            Height          =   255
            Left            =   2520
            TabIndex        =   27
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame frNroDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   9285
         Begin VB.TextBox txtDes_TipDoc 
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Top             =   375
            Width           =   1905
         End
         Begin VB.TextBox txtCod_TipDoc 
            Height          =   285
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   19
            Top             =   375
            Width           =   480
         End
         Begin VB.TextBox txtSer_Docum 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   18
            Top             =   375
            Width           =   540
         End
         Begin VB.TextBox txtNum_Docum 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            MaxLength       =   8
            TabIndex        =   17
            Top             =   375
            Width           =   1080
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Serie :"
            Height          =   255
            Left            =   4080
            TabIndex        =   23
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Número :"
            Height          =   195
            Left            =   5400
            TabIndex        =   22
            Tag             =   "Number"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tipo Doc :"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   390
            Width           =   855
         End
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   10440
         TabIndex        =   42
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71696385
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   10440
         TabIndex        =   43
         Top             =   840
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71696385
         CurrentDate     =   37543
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha de Emision"
         Height          =   195
         Left            =   10440
         TabIndex        =   46
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   9720
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desde :"
         Height          =   255
         Left            =   9720
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame fraFecCobRepro 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha Cobranza Reprogramada"
      Height          =   1650
      Left            =   4320
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   3555
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   600
         TabIndex        =   14
         Top             =   960
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowFactVentas.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin NumBoxProject.NumBox txtFecCobRepro 
         Height          =   330
         Left            =   960
         TabIndex        =   47
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   0
         MaskLen         =   20
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
   End
   Begin VB.Frame fraReactivarFactura 
      Caption         =   "Aceptar"
      Height          =   2295
      Left            =   2880
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtCod_TipDoc2 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   "FA"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtSer_Docum2 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtNum_Docum2 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo Doc"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Serie"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Número"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Fra_FechaRecepcion 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha Recepcion"
      Height          =   1650
      Left            =   7320
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   3555
      Begin FunctionsButtons.FunctButt FunctButt4 
         Height          =   510
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowFactVentas.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin NumBoxProject.NumBox txtFecrecepcion 
         Height          =   330
         Left            =   1200
         TabIndex        =   48
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   0
         MaskLen         =   20
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   7095
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   12515
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmShowFactVentas.frx":012E
      FormatStyle(2)  =   "frmShowFactVentas.frx":0266
      FormatStyle(3)  =   "frmShowFactVentas.frx":0316
      FormatStyle(4)  =   "frmShowFactVentas.frx":03CA
      FormatStyle(5)  =   "frmShowFactVentas.frx":04A2
      FormatStyle(6)  =   "frmShowFactVentas.frx":055A
      FormatStyle(7)  =   "frmShowFactVentas.frx":063A
      ImageCount      =   0
      PrinterProperties=   "frmShowFactVentas.frx":065A
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   6435
      Left            =   12720
      TabIndex        =   13
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   11351
      Custom          =   $"frmShowFactVentas.frx":0832
      Orientacion     =   1
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   493
      ControlSeparator=   0
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6435
      Top             =   7920
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowFactVentas"
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
Dim strOpcion As String
Public CODIGO As String, Descripcion As String, strCod_Anxo As String, TipoAdd As String
Public SNum_Corre As String
Public SImp_Total As Double
Dim iLin As Integer
Dim lvSW As Boolean
Private stipdoc_venta As String
Dim factura_guia As String
Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
  dtpFecEmiFin.Value = dtpFecEmiIni.Value
End Sub

Private Sub Form_Load()
  lvSW = True
  dtpFecEmiIni.Value = Date
  dtpFecEmiFin.Value = Date

  'FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR/BUSCAR"
  'FunctButt3.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
  
  strOpcion = "C"
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  optNroDoc.Value = True
  Call optNroDoc_Click

End Sub
Private Sub cmdBuscar_Click()
  BUSCAR
End Sub
Sub BUSCAR()
On Error GoTo dprDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSQL = "Ventas_Muestra_Doc_Ventas_rollo '" & strOpcion & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & "C" & "','" & strCod_Anxo & "','" & txtAno & "','" & txtMes & "','','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & "','" & vusu & "','','" & Trim(txtPartidaBusqueda.Text) & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
GridEX1.Columns("Cod_TipDoc").Width = 375
GridEX1.Columns("Cod_TipDoc").Caption = "Tip"
GridEX1.Columns("Serie").Width = 525
GridEX1.Columns("Serie").Caption = "Serie"
GridEX1.Columns("Nro_Doc").Width = 810
GridEX1.Columns("Nro_Doc").Caption = "Nro_Doc"
GridEX1.Columns("Anexo").Width = 2865
GridEX1.Columns("Anexo").Caption = "Anexo"
GridEX1.Columns("Ruc").Width = 1410
GridEX1.Columns("Ruc").Caption = "Ruc"
GridEX1.Columns("Moneda").Width = 705
GridEX1.Columns("Moneda").Caption = "Moneda"
GridEX1.Columns("Imp_Neto").Width = 825
GridEX1.Columns("Imp_Neto").Caption = "Imp Neto"
GridEX1.Columns("Imp_Igv").Width = 705
GridEX1.Columns("Imp_Igv").Caption = "Imp Igv"
GridEX1.Columns("Imp_Gastos_Financieros").Caption = "Gastos Financieros"
GridEX1.Columns("Imp_Gastos_Financieros").Width = 990
GridEX1.Columns("Imp_Total").Width = 840
GridEX1.Columns("Imp_Total").Caption = "Imp Total"
GridEX1.Columns("Imp_Otros").Width = 870
GridEX1.Columns("Imp_Otros").Caption = "Imp Otros"
GridEX1.Columns("Emision").Width = 945
GridEX1.Columns("Emision").Caption = "Emision"
GridEX1.Columns("Registro").Width = 945
GridEX1.Columns("Registro").Caption = "Registro"
GridEX1.Columns("Vencimiento").Width = 945
GridEX1.Columns("Vencimiento").Caption = "Vencimiento"
GridEX1.Columns("Cancelado").Width = 1500
GridEX1.Columns("Cancelado").Caption = "Cancelado"
GridEX1.Columns("Ano_Registro").Width = 1095
GridEX1.Columns("Ano_Registro").Caption = "Ano_Registro"
GridEX1.Columns("Mes_Registro").Width = 1110
GridEX1.Columns("Mes_Registro").Caption = "Mes_Registro"
GridEX1.Columns("Num_Registro").Width = 1140
GridEX1.Columns("Num_Registro").Caption = "Num_Registro"

GridEX1.Columns("Num_Dua").Width = 1140
GridEX1.Columns("Num_Dua").Caption = "Num_Dua"
GridEX1.Columns("Fec_NumeracionDua").Width = 1140
GridEX1.Columns("Fec_NumeracionDua").Caption = "Fec_NumeracionDua"
GridEX1.Columns("Fec_EmbarqueReal").Width = 1140
GridEX1.Columns("Fec_EmbarqueReal").Caption = "Fec_EmbarqueReal"
GridEX1.Columns("Imp_FOB_Dol_Dua").Width = 1140
GridEX1.Columns("Imp_FOB_Dol_Dua").Caption = "Imp_FOB_Dol_Dua"

GridEX1.Columns("Emision").Format = "dd/mm/yyyy"
GridEX1.Columns("Registro").Format = "dd/mm/yyyy"
GridEX1.Columns("Vencimiento").Format = "dd/mm/yyyy"
GridEX1.Columns("Cancelado").Format = "dd/mm/yyyy"
GridEX1.Columns("Fec_NumeracionDua").Format = "dd/mm/yyyy"
GridEX1.Columns("Fec_EmbarqueReal").Format = "dd/mm/yyyy"


GridEX1.Columns("Transmision").Width = 0
GridEX1.Columns("Despacho").Width = 0
GridEX1.Columns("Impresion").Width = 0
'GridEX1.Columns("Imp_Descuento").Width = 1140
'GridEX1.Columns("Imp_descuento").Caption = "IMP_DESCUENTO"


GridEX1.ContinuousScroll = True

Exit Sub

dprDepurar:

errores err.Number
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub



Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo HandlerError

    Dim Msg As Variant
    
    Select Case ActionName
    
    Case "BUSCAR"
      BUSCAR
    Case "ADICIONAR"
      With frmAdicionaDocumVentas
        .StrOption = "I"
        
        .Show 1
        If .strNum_Corre <> "" Then
        
          'optCorrelativo = True
                    BUSCAR
drpsiguiente:

'          If GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "NC" Or GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "ND" Then
'            Load frmAdicionaDetalleDocumAsigNotas
'            With frmAdicionaDetalleDocumAsigNotas
'              .Caption = "Adicion " + GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) & " Nro " & GridEX1.Value(GridEX1.Columns("Serie").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nro_Doc").Index) & " Del Cliente " & GridEX1.Value(GridEX1.Columns("Anexo").Index)
'              '.strNum_Corre_Ori = txtNum_Corre
'              .Show 1
'            End With
'          Else
'            Load frmAdicionaDetalleDocum
'            With frmAdicionaDetalleDocum
'              .Caption = "Adicion " + GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) & " Nro " & GridEX1.Value(GridEX1.Columns("Serie").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nro_Doc").Index) & " Del Cliente " & GridEX1.Value(GridEX1.Columns("Anexo").Index)
'              '.strNum_Corre_Detalle = txtNum_Corre
'              .IntSencuencia = 0
'              .strOption = "I"
'              .Show 1
'            End With
'          End If
'
'          If frmAdicionaDetalleDocum.IntSencuencia <> 0 Then
'            GoTo drpsiguiente
'          Else
'            Call FunctButt1_ActionClick(0, 0, "VERDETALLE")
'          End If
      
      End If
      Set frmAdicionaDocumVentas = Nothing
      End With
      BUSCAR
    Case "MODIFICA"
    
      If GridEX1.RowCount = 0 Then Exit Sub
      If GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "FA" Or GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "BV" Or GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "TK" Then
            With FrmModificaDocumentoVentas  'frmAdicionaDocumVentas
              .StrOption = "U"
              If DevuelveCampo("Select dbo.ventas_Valida_Documento_Manuales_Cabrezera('" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "')", cConnect) <> "" Then
                MsgBox "Documento cerrado.", vbInformation, "AVISO"
                FrmModificaDocumentoVentas.frMain.Enabled = False
                'Call FrmModificaDocumentoVentas.adicionarProductoDesdeDetallefactura(GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("cod_tipdoc").Index), GridEX1.Value(GridEX1.Columns("ser_docum").Index), GridEX1.Value(GridEX1.Columns("num_docum_ventas").Index))
                'frmAdicionaDocumVentas.frExportacion.Enabled = False
                FrmModificaDocumentoVentas.FunctButt1.Visible = False
                .StrOption = "A"
              End If
              
              .Caption = "Modifica Docum Ventas"
              
              Carga_Data
              .strNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
              '.adicionarProductoDesdeDetallefactura (GridEX1.Value(GridEX1.Columns("cod_tipdoc").Index), GridEX1.Value(GridEX1.Columns("ser_docum").Index), GridEX1.Value(GridEX1.Columns("num_docum_ventas").Index))
              Call FrmModificaDocumentoVentas.adicionarProductoDesdeDetallefactura("", "", "", GridEX1.Value(GridEX1.Columns("Num_Corre").Index))
              .Show 1

              Set FrmModificaDocumentoVentas = Nothing
              If .strNum_Corre <> "" Then
                'optCorrelativo = True
                'txtNum_Corre = .strNum_Corre
                BUSCAR
              End If

            End With
      End If
      
      If GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "NC" Or GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "ND" Then
            With frmDocumentoVentaAnterior
              .StrOption = "U"
              If DevuelveCampo("Select dbo.ventas_Valida_Documento_Manuales_Cabrezera('" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "')", cConnect) <> "" Then
                MsgBox "Documento cerrado.", vbInformation, "AVISO"
                frmDocumentoVentaAnterior.frMain.Enabled = False
                frmDocumentoVentaAnterior.frExportacion.Enabled = False
                frmDocumentoVentaAnterior.FunctButt1.Visible = False
                .StrOption = "A"
              End If
              .Caption = "Modifica Docum Ventas"
              Carga_Data_ANTERIOR
              .strNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
              .Show 1
              Set frmDocumentoVentaAnterior = Nothing
              If .strNum_Corre <> "" Then
                  'optCorrelativo = True
                  'txtNum_Corre = .strNum_Corre
                  BUSCAR
              End If
            End With
            BUSCAR
      End If
      
     BUSCAR
     
    Case "VERDETALLE"
      If GridEX1.RowCount = 0 Then Exit Sub
      Load frmMuestraDetalleDocumVentas
      With frmMuestraDetalleDocumVentas
        .Caption = GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) & " Nro " & GridEX1.Value(GridEX1.Columns("Serie").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nro_Doc").Index) & " Del Cliente " & GridEX1.Value(GridEX1.Columns("Anexo").Index)
        .StrSQL = "Ventas_Muestra_Detalle_Factura_Items '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"
        .strCod_TipDoc = GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)
        .Num_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
        .BUSCAR
        .Show 1
        BUSCAR
      End With
      BUSCAR

    Case "GENERABAUCHER"
         If GridEX1.RowCount = 0 Then Exit Sub
        Genera_Voucher
        BUSCAR
    Case "VERVOUCHER"
     If GridEX1.RowCount = 0 Then Exit Sub
     MuestraVoucher2
        Load frmShowCN_DocumVoucher_Ventas
        Set frmShowCN_DocumVoucher_Ventas.oParent = Me
        frmShowCN_DocumVoucher_Ventas.SNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
        frmShowCN_DocumVoucher_Ventas.Caption = frmShowCN_DocumVoucher_Ventas.Caption & " Correlativo : " & frmShowCN_DocumVoucher_Ventas.SNum_Corre & " Nº Registro : " & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & " !!! "
        frmShowCN_DocumVoucher_Ventas.BUSCAR
        frmShowCN_DocumVoucher_Ventas.Show vbModal
        Set frmShowCN_DocumVoucher_Ventas = Nothing
     BUSCAR
    Case "RAIMPRECION"
        If GridEX1.RowCount = 0 Then Exit Sub
    BUSCAR
        'Imprimir GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("Imp_Total").Index), False, GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)
    Case "ANULAR"
        If GridEX1.RowCount = 0 Then Exit Sub
        stipdoc_venta = UCase(GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index))
        factura_guia = DevuelveCampo("select isnull(FLG_FACT_MOV_ALM ,'') from cn_ventas where num_corre= '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'", cConnect)
        
        If MsgBox("Esta Seguro de Anular este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
        
           If Trim(factura_guia) = "" Or stipdoc_venta = "NC" Or stipdoc_venta = "ND" Then
                ExecuteCommandSQL cConnect, "Ventas_Man_Anula_Docum '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
           End If
          If stipdoc_venta = "FA" Or stipdoc_venta = "BV" Or stipdoc_venta = "TK" Then
                '''factura guia elimina el movimiento de almacen
                If factura_guia = "S" Then
                     ExecuteCommandSQL cConnect, "VENTAS_MAN_ANULA_DOCUM_ROLLO '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
                End If
                '''no elimina los movimiento de almacen
                If factura_guia = "N" Then
                     ExecuteCommandSQL cConnect, "VENTAS_MAN_ANULA_DOCUM_ROLLO_GUIA'" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
                End If
          End If
        End If
        BUSCAR
    Case "REVIERTEDOCUM"
        
        'If GridEX1.RowCount = 0 Then Exit Sub
        'If MsgBox("Esta Seguro de Revertir este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
        '   ExecuteCommandSQL cConnect, "Ventas_Revierte_Docum_rollo '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
        '   BUSCAR
        'End If
        'BUSCAR
        If GridEX1.RowCount = 0 Then Exit Sub
        'Dim factura_guia As String
        stipdoc_venta = UCase(GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index))
        factura_guia = DevuelveCampo("select isnull(FLG_FACT_MOV_ALM,'') from cn_ventas where num_corre= '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'", cConnect)
        
        If MsgBox("Esta Seguro de Revertir este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
            If Trim(factura_guia) = "" Or stipdoc_venta = "NC" Or stipdoc_venta = "ND" Then
                ExecuteCommandSQL cConnect, "Ventas_Revierte_Docum '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
            End If
            If stipdoc_venta = "FA" Or stipdoc_venta = "BV" Or stipdoc_venta = "TK" Then
                If factura_guia = "S" Then
                    ExecuteCommandSQL cConnect, "VENTAS_MAN_REVIERTE_DOCUM_ROLLO_ALMACEN '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
                                                 
                End If
                If factura_guia = "N" Then
                    ExecuteCommandSQL cConnect, "VENTAS_REVIERTE_DOCUM_ROLLO '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
                End If
            End If
        End If
        BUSCAR
    Case "IMPRESIONES"
        If GridEX1.RowCount = 0 Then Exit Sub
        
            SNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
            SImp_Total = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
            frmImpresionesFacturas.SNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
            frmImpresionesFacturas.SImp_Total = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
            frmImpresionesFacturas.SCod_TipDoc = GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)
            Call Factura_sa(GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index), GridEX1.Value(GridEX1.Columns("Imp_Total").Index))

'            If GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "FA" Or GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index) = "BV" Then
'                IMPRIMIR_DOCSALIDA
'            Else
'                frmImpresionesFacturas.Show vbModal
'                Set frmImpresionesFacturas = Nothing
'            End If
             BUSCAR
        
    Case "ADICIONANCND"
        With frmDocumentoVentaAnterior
            .StrOption = "I"
            .Show 1
        End With
        Set frmDocumentoVentaAnterior = Nothing
        BUSCAR
      
    Case "DESPACHO_EXTEMPORANE"
     
        If GridEX1.RowCount = 0 Then Exit Sub
        'Load frmConfirmacionDespacho
        frmConfirmacionDespacho.Cod_TipDoc = GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)
        frmConfirmacionDespacho.Serie = GridEX1.Value(GridEX1.Columns("Serie").Index)
        frmConfirmacionDespacho.Nro_doc = GridEX1.Value(GridEX1.Columns("Nro_Doc").Index)
        'frmConfirmacionDespacho.Valor = GridEX1.Value(GridEX1.Columns("Despacho").Index)
        
        Set frmConfirmacionDespacho.oParent = Me
        frmConfirmacionDespacho.Show vbModal
        Set frmConfirmacionDespacho = Nothing
        BUSCAR
        
     Case "REVIERTEIMPRESION"
     If GridEX1.RowCount = 0 Then Exit Sub
     
     If MsgBox("Esta Seguro de Revertir este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
                ExecuteCommandSQL cConnect, "VENTAS_REVIERTE_IMPRESION '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
                Call MsgBox("La Reversion de Impresion se realizo con exito", vbInformation, "Mensaje")
     End If
     
     Case "LDP/DDP"
        If GridEX1.RowCount = 0 Then Exit Sub
        frmCompletarImportesLDPDDP.strNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
        frmCompletarImportesLDPDDP.txtFlete = GridEX1.Value(GridEX1.Columns("Imp_flete").Index)
        frmCompletarImportesLDPDDP.txtDesaduanaje = GridEX1.Value(GridEX1.Columns("imp_desaduanaje").Index)
        frmCompletarImportesLDPDDP.txtTransporte = GridEX1.Value(GridEX1.Columns("imp_transporte_pais_destino").Index)
        frmCompletarImportesLDPDDP.txtFob = GridEX1.Value(GridEX1.Columns("Imp_FOB").Index)
        frmCompletarImportesLDPDDP.txtCif = GridEX1.Value(GridEX1.Columns("Imp_CIF").Index)
        frmCompletarImportesLDPDDP.txtLdp = GridEX1.Value(GridEX1.Columns("Imp_LDP").Index)
        frmCompletarImportesLDPDDP.txtDdp = GridEX1.Value(GridEX1.Columns("Imp_DDP").Index)
       
        frmCompletarImportesLDPDDP.Show vbModal
        Set frmCompletarImportesLDPDDP = Nothing
    BUSCAR
    Case "REDONDEAR"
     Dim frmRe As New frmRedondeaImporte
     If GridEX1.RowCount = 0 Then Exit Sub
       With frmRe
         .txtImporteNeto = GridEX1.Value(GridEX1.Columns("Imp_Neto").Index)
         .txtImporteIgv = GridEX1.Value(GridEX1.Columns("Imp_Igv").Index)
         .txtImporteTotal = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
         .txtImpTotalActual = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
         .txtImporteGastosFinan = GridEX1.Value(GridEX1.Columns("Imp_Gastos_Financieros").Index)
         .txtImporteDscto = GridEX1.Value(GridEX1.Columns("Imp_descuento").Index)
         .ValorActualIGV = GridEX1.Value(GridEX1.Columns("Imp_Igv").Index)
         .ValorActualImporteNeto = GridEX1.Value(GridEX1.Columns("Imp_Neto").Index)
         .ValorActualImporteTotalR = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
         .ValorActualImporteTotal = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
         .txtImporteOtros = GridEX1.Value(GridEX1.Columns("Imp_Otros").Index)
         .porcIGV = GridEX1.Value(GridEX1.Columns("porc_igv").Index)
         Set .grilla = GridEX1
         .Show 1
         BUSCAR
       End With
       
       BUSCAR
    Case "IMPEXP"
        If GridEX1.RowCount = 0 Then Exit Sub
       ' Imprimir_Exp GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index), GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
        Imprimir_Exp GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index), GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
    BUSCAR
    Case "IMPORTEFOBDUA"
        If GridEX1.RowCount = 0 Then Exit Sub
        Dim frmImpFDua As New frmImporteFobDua
        frmImpFDua.numCorre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
        frmImpFDua.txtDua.Text = GridEX1.Value(GridEX1.Columns("Num_Dua").Index)
        'frmImpFDua.txtFec_Numeracion.Text = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Fec_NumeracionDua").Index)), "", GridEX1.Value(GridEX1.Columns("Fec_NumeracionDua").Index))
       ' frmImpFDua.txtFec_Embarque.Text = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Fec_EmbarqueReal").Index)), "", GridEX1.Value(GridEX1.Columns("Fec_EmbarqueReal").Index))
       '  frmImpFDua.txtImp_FOB_Dol_Dua.Text = GridEX1.Value(GridEX1.Columns("Imp_FOB_Dol_Dua").Index)
        frmImpFDua.Show 1
        Set frmImpFDua = Nothing
  BUSCAR
    Case "GENERAINFOCONT"
        If GridEX1.RowCount = 0 Then Exit Sub
        GeneraInfoContable
   BUSCAR
    Case "RINFCONTABLE"
        If GridEX1.RowCount = 0 Then Exit Sub
        If MsgBox("Esta Seguro de Revertir Ifx Contable de este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
           ExecuteCommandSQL cConnect, "CN_REVIERTE_ASIENTO_VENTAS'" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"
           Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
           BUSCAR
        End If
    BUSCAR
    Case "FECDOC"
         
         If GridEX1.RowCount = 0 Then Exit Sub
         frm_FecEnvDoc.Cod_TipDoc = GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)
         frm_FecEnvDoc.Serie = GridEX1.Value(GridEX1.Columns("Serie").Index)
         frm_FecEnvDoc.Nro_doc = GridEX1.Value(GridEX1.Columns("Nro_Doc").Index)
         frm_FecEnvDoc.DTPFecha.Value = GridEX1.Value(GridEX1.Columns("Fec_Envio_Documentos_Cobranza").Index)
         Set frm_FecEnvDoc.oParent = Me
         frm_FecEnvDoc.Show vbModal
         Set frm_FecEnvDoc = Nothing
         BUSCAR
         
    Case "FECCOBREPRO"
        If GridEX1.RowCount = 0 Then Exit Sub
        GridEX1.Enabled = False
        'txtFecCobRepro.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_Cobranza_Reprogramada").Index), vbString)
       
        fraFecCobRepro.Visible = True
    Case "APLICANCALGO"
        If GridEX1.RowCount = 0 Then Exit Sub
        Call AplicaNotaCreditoAlgolimsa
        
    Case "VERCOBROS"
        If GridEX1.RowCount = 0 Then Exit Sub
        Load FrmVer_Cobros
        FrmVer_Cobros.SNum_Corre = Trim(FixNulos(GridEX1.Value(GridEX1.Columns("NUM_CORRE").Index), vbString))
        FrmVer_Cobros.StrSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & Trim(GridEX1.Value(GridEX1.Columns("Num_Corre").Index)) & "'"
        FrmVer_Cobros.BUSCAR
        FrmVer_Cobros.Show vbModal
        Set FrmVer_Cobros = Nothing
        
    Case "ELIMFANULADA"
        Me.fraReactivarFactura.Visible = True
        Me.txtSer_Docum2.SetFocus
    
    Case "ACTNFOB"
        If GridEX1.RowCount > 0 Then
            actualizarNoFob
            BUSCAR
        Else
            MsgBox "Seleccione un Registro", vbExclamation, "Mensaje"
            Exit Sub
        End If
    Case "SALIR"
    Unload Me
    End Select
Exit Sub
Resume
HandlerError:
  errores err.Number
End Sub


Sub actualizarNoFob()
On Error GoTo errGrabar
Dim numCorre As String
Dim vMessage As Variant
Dim StrSQL As String

numCorre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
        
vMessage = MsgBox("Desea Actualizar al estado NO FOB", 48 + 4, "Actualizar Factura")
    If vMessage = vbYes Then
        StrSQL = "CN_VENTAS_ACTUALIZA_IMPORTE_NO_FOB '" & numCorre & "'"
        Call ExecuteCommandSQL(cConnect, StrSQL)
        MsgBox "Transaccion Realizada con Exito", vbInformation, "Mensaje"
        Exit Sub
    End If
Exit Sub
errGrabar:
    MsgBox err.Description, vbCritical, "cerrarCarta"
End Sub

Sub AplicaNotaCreditoAlgolimsa()
Dim StrSQL As String
On Error GoTo lblError
    If MsgBox("¿Seguro que desea aplicar?", vbYesNo + vbQuestion, "Mensaje del Sistema") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    StrSQL = "EXEC VENTAS_APLICACION_NC_ALGOLIMSA '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"
    Call ExecuteCommandSQL(cConnect, StrSQL)
    Call BUSCAR
    Screen.MousePointer = vbDefault
Exit Sub
lblError:
    MsgBox err.Description, vbCritical, "Mensaje del sistema"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Imprimir_Exp(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String, dbImp_Total As Double)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

   ' sFormato_Invoice = DevuelveCampo("SELECT FORMATO_INVOICE FROM TG_CLIENTE WHERE COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
   ' Select Case sCod_Tipdoc
    '    Case "FA"
            oo.Workbooks.Open vRuta & "\Invoice03_1.XLT" ' & sFormato_Invoice & ".XLT"
   ' End Select
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "reporte", cConnect, SNum_Corre, UCase(EnLetras(Trim(CStr(dbImp_Total))))
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub


Private Function ifValidaDoc() As Boolean

Dim strMsg As String

strMsg = DevuelveCampo("Select dbo.ventas_Valida_Documento_Manuales('" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "')", cConnect)
If strMsg <> "" Then
  MsgBox strMsg, vbInformation, "AVISO"
  ifValidaDoc = False
  Exit Function
End If

ifValidaDoc = True

End Function

Sub Carga_Data_ANTERIOR()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set RS = CargarRecordSetDesconectado("Ventas_Up_Man 'V','" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'", cConnect)

With RS
  If Not (.BOF Or .EOF) Then
    With frmDocumentoVentaAnterior
    
      .txtCod_TipDoc = RS!Cod_TipDoc
      .txtDes_TipDoc = RS!Des_TipDoc
      .txtCod_TipVenta = RS!Cod_Tipo_Venta
      .txtDes_TipVenta = RS!Des_Tipo_Vent
      .Cambio_FR
      .txtSer_Docum = RS!ser_docum
      .txtNum_Docum = RS!num_docum_ventas
      .strCod_Anxo = RS!Cod_Anxo
      .txtDes_TipAne = RS!Des_Anexo
      .txtNum_Ruc = RS!Num_Ruc
      '.inpFec_EmiDoc.Text = RS!Fec_EmiDoc
      '.InpFec_RegDoc.Text = RS!Fec_RegDoc
      .TxtTipo_Cambio.Text = RS!Tipo_Cambio
      .txtCod_Moneda = RS!Cod_Moneda
      .txtDes_Moneda = RS!Nom_Moneda
      .txtCod_ConPag = RS!Cod_CondVent
      .txtDes_ConPag = RS!Des_CondVent
      .txtNro_Guias = RS!Guias
      .txtNro_Ordener = RS!Pedidos
      .txtNro_DocInter = RS!Partes
      .Imp_Gastos_Finacieros.Text = RS!Imp_Gastos_Financieros
      .Imp_Otros.Text = RS!Imp_Otros
      .txtGlosa = RS!Glosa
      
      .txtCod_TipDoc.Enabled = False
      .txtDes_TipDoc.Enabled = False
      .txtSer_Docum.Enabled = False
      .txtNum_Docum.Enabled = False
      '.txtCod_TipVenta.Enabled = False
      '.txtDes_TipVenta.Enabled = False
      
      .chkExportacion.Value = IIf(RS!Flg_Exportacion <> "S", 0, 1)
      .chkFlete.Value = IIf(RS!Flg_Inc_Flete_Export <> "S", 0, 1)
      .chkSeguro.Value = IIf(RS!Flg_Inc_Seguro_Export <> "S", 0, 1)
      .chkDetraccion.Value = IIf(RS!Flg_Retencion_IGV <> "S", 0, 1)
      .chkExonerado.Value = IIf(RS!Flg_Exonerado_IGV <> "S", 0, 1)
      
      .txtEmbarque_Cod = RS!Tip_Embarque
      .txtEmbarque_Des = RS!Des_TipEmbarque
      
      
      'If GridEX1.Value(GridEX1.Columns("Transmision").Index) <> "P" Or GridEX1.Value(GridEX1.Columns("Impresion").Index) <> "N" Then .frMain.Enabled = False
      'If .chkExportacion.Value Then .frMain.Enabled = True
      
      .Imp_Flete.Text = RS!Imp_Flete
      .txtReferencia = RS!Glosa_Documento_Referencia
      
      .txtObservacion = RS!Observacion
      .txtCod_Destino = RS!Cod_Destino
      .txtDes_Destino = RS!Des_Destino
      .txtShip_Date.Text = RS!Ship_Date
      .txtPeso_Bruto.Text = RS!Peso_Bruto
      .txtPeso_Neto.Text = RS!Peso_Neto
      .txtImp_Seguro.Text = RS!Imp_Seguro

      .txtCod_NotaAbono = RS!Cod_Mot_Nota
      .txtDes_NotaAbono = RS!Des_MotAbono

      .txtDua.Text = RS!Num_Dua
      .txtFec_Numeracion.Text = RS!Fec_NumeracionDua
      .txtFec_Embarque.Text = RS!Fec_EmbarqueReal
      .txtImp_FOB_Dol_Dua.Text = RS!Imp_FOB_Dol_Dua
      .txtcajas.Text = RS!Num_Bultos
      
    End With
  End If
End With

End Sub
Sub Carga_Data()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set RS = CargarRecordSetDesconectado("Ventas_Up_Man 'V','" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'", cConnect)

With RS
  If Not (.BOF Or .EOF) Then
    With FrmModificaDocumentoVentas 'frmAdicionaDocumVentas
    
      .txtCod_TipDoc = RS!Cod_TipDoc
      .txtDes_TipDoc = RS!Des_TipDoc
      .txtCod_TipVenta = RS!Cod_Tipo_Venta
      .txtDes_TipVenta = RS!Des_Tipo_Vent
      .Cambio_FR
      .txtSer_Docum = RS!ser_docum
      .txtNum_Docum = RS!num_docum_ventas
      .strCod_Anxo = RS!Cod_Anxo
      .txtDes_TipAne = RS!Des_Anexo
      .txtNum_Ruc = RS!Num_Ruc
      '.inpFec_EmiDoc.Text = RS!Fec_EmiDoc
      '.InpFec_RegDoc.Text = RS!Fec_RegDoc
      .TxtTipo_Cambio.Text = RS!Tipo_Cambio
      .txtCod_Moneda = RS!Cod_Moneda
      .txtDes_Moneda = RS!Nom_Moneda
      .txtCod_ConPag = RS!Cod_CondVent
      .txtDes_ConPag = RS!Des_CondVent
      '.txtNro_Guias = RS!Guias
      '.txtNro_Ordener = RS!Pedidos
      '.txtNro_DocInter = RS!Partes
      '.Imp_Gastos_Finacieros.Text = RS!Imp_Gastos_Financieros
      '.Imp_Otros.Text = RS!Imp_Otros
      '.txtGlosa = RS!Glosa
      
      .txtCod_TipDoc.Enabled = False
      .txtDes_TipDoc.Enabled = False
      .txtSer_Docum.Enabled = False
      .txtNum_Docum.Enabled = False
      '.txtCod_TipVenta.Enabled = False
      '.txtDes_TipVenta.Enabled = False
      
      '.chkExportacion.Value = IIf(RS!Flg_Exportacion <> "S", 0, 1)
      '.chkFlete.Value = IIf(RS!Flg_Inc_Flete_Export <> "S", 0, 1)
      '.chkSeguro.Value = IIf(RS!Flg_Inc_Seguro_Export <> "S", 0, 1)
      '.chkDetraccion.Value = IIf(RS!Flg_Retencion_IGV <> "S", 0, 1)
      '.chkExonerado.Value = IIf(RS!Flg_Exonerado_IGV <> "S", 0, 1)
    
    End With
  End If
End With

End Sub







Private Sub Genera_Voucher()
On Error GoTo fin
Dim sTit As String
Dim sAccion As String, StrSQL As String

sAccion = "D"
   sTit = "Generar Voucher De Ventas"
    
   If MsgBox("Genera Voucher De Ventas...?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
    StrSQL = "EXEC CN_GENERA_VOUCHER_VENTAS '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & vusu & "'"
    
    
    ExecuteCommandSQL cConnect, StrSQL
  BUSCAR
    
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
        Call ActualizaFechaCobranzaReprogramada
Case "CANCELAR"
        fraFecCobRepro.Visible = False
        GridEX1.Enabled = True
End Select
End Sub
Private Sub ActualizaFechaCobranzaReprogramada()
On Error GoTo errores
Dim sSQL As String
Dim sFlg_Pendalidad As String

'If txtFecCobRepro.Text = "" Then
'    sSQL = "CN_VENTAS_ACTUALIZA_FEC_COBRANZA_REPROGRAMADA '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "',null"
'Else
'    sSQL = "CN_VENTAS_ACTUALIZA_FEC_COBRANZA_REPROGRAMADA '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & txtFecCobRepro.Text & "'"
'End If
'
'ExecuteCommandSQL cConnect, sSQL
'GridEX1.Enabled = True
'Me.fraFecCobRepro.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub


Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

    Select Case ActionName
        Case "MOVTELATENIDA"
           Reporte
        Case "RECEPCION"
         If GridEX1.RowCount = 0 Then Exit Sub
            GridEX1.Enabled = False
            txtFecrecepcion.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fecha_Recepcion").Index), vbString)
       
            Fra_FechaRecepcion.Visible = True
    End Select
End Sub

Private Sub Reporte()
   On Error GoTo SALTO_ERROR
    Dim oRs As New Recordset
    Dim StrSQL As String
    If Trim(GridEX1.Value(GridEX1.Columns("num_corre").Index)) <> "" Then
        StrSQL = "VENTAS_MUESTRA_MOVS_TELA_TENIDA_SEGUN_FACTURA '" & Trim(GridEX1.Value(GridEX1.Columns("num_corre").Index)) & "'"
    
        Set oRs = CargarRecordSetDesconectado(StrSQL, cConnect)
        If oRs.RecordCount = 0 Then
            MsgBox "No se han encontrado datos para la impresión.....", vbExclamation
            Exit Sub
        End If
        
        Dim oo As Object
        Dim sRutaLogo As String, sTitulo As String
        
        Set oo = CreateObject("excel.application")
        StrSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
        sRutaLogo = DevuelveCampo(StrSQL, cConnect)
        oo.Workbooks.Open vRuta & "\rptMovTelaTenida.XLT"
        oo.Visible = True
        oo.DisplayAlerts = False
        
        oo.run "reporte", sRutaLogo, oRs, Trim(GridEX1.Value(GridEX1.Columns("num_corre").Index))
        
        Set oo = Nothing
    Else
        MsgBox "Seleccione un Documento"
    End If
    
    Exit Sub

SALTO_ERROR:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub FunctButt4_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
        Call ActualizaFechaRecepcion
Case "CANCELAR"
        Fra_FechaRecepcion.Visible = False
        GridEX1.Enabled = True
End Select
End Sub

Private Sub optAnoMes_Click()
  limpiar
  frAnoMes.Visible = True
  strOpcion = "A"
  txtAno.SetFocus
End Sub

Private Sub optcliente_Click()
  limpiar
  frCliente.Visible = True
  strOpcion = "C"
  txtNum_Ruc.SetFocus
End Sub
Sub limpiar()

  frCliente.Visible = False
  frAnoMes.Visible = False
  frNroDoc.Visible = False
  frPartida.Visible = False

  txtNum_Ruc.Text = ""
  txtDes_Anexo.Text = ""
  txtAno.Text = ""
  txtMes.Text = ""
  txtCod_TipDoc.Text = ""
  txtDes_TipDoc.Text = ""
  txtSer_Docum.Text = ""
  txtNum_Docum.Text = ""
  txtPartidaBusqueda.Text = ""
  
End Sub

Private Sub optNroDoc_Click()
  limpiar
  frNroDoc.Visible = True
  strOpcion = "N"
  'txtCod_TipDoc.SetFocus
End Sub

Private Sub optPartida_Click()
  limpiar
  frPartida.Visible = True
  strOpcion = "P"
  txtPartidaBusqueda.SetFocus
End Sub


Private Sub txtAno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtCod_TipAnxo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAnxo, txtDes_TipDoc, 1, Me)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
  End If
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
  'If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='" & Trim(txtCod_TipAnxo.Text) & "' and ", txtNum_Ruc, txtDes_Anexo, 2, Me)
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2, Me)
End Sub

Private Sub txtFecCobRepro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    FunctButt2.SetFocus
End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub


Private Sub txtNum_Corre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Docum_LostFocus()
  txtNum_Docum = Format(txtNum_Docum, "00000000")
  FunctButt1.SetFocus
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  'Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='" & Trim(txtCod_TipAnxo.Text) & "' and ", txtNum_Ruc, txtDes_Anexo, 1, Me)
  SendKeys "{TAB}"
 End If
End Sub
Private Sub txtPartidaBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys "{TAB}"
FunctButt1.SetFocus
End If
End Sub

Private Sub txtPartidaBusqueda_LostFocus()
txtPartidaBusqueda = Format(txtPartidaBusqueda, "00000")
'FunctButt1.SetFocus
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'SendKeys "{TAB}"
     txtNum_Docum.SetFocus
  End If
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub

Private Sub MuestraVoucher2()
On Error GoTo errx
Dim sSQL As String
Dim rsAsientos As Object
Set rsAsientos = CreateObject("ADODB.Recordset")


If GridEX1.RowCount = 0 Then Exit Sub

  If RTrim(GridEX1.Value(GridEX1.Columns("num_registro").Index)) = "" Then
    MsgBox "No se le ha Generado Voucher", vbInformation, "AVISO"
    Exit Sub
  End If

  Load frmShowVoucher
  frmShowVoucher.sCod_TipoDiario = RTrim(DevuelveCampo("select Cod_TipodiarioVentas  from cn_control ", cConnect))
  frmShowVoucher.sano = RTrim(GridEX1.Value(GridEX1.Columns("Ano_Registro").Index))
  frmShowVoucher.smes = RTrim(GridEX1.Value(GridEX1.Columns("Mes_registro").Index))
  frmShowVoucher.lNum_Registro = RTrim(GridEX1.Value(GridEX1.Columns("Num_Registro").Index))
  frmShowVoucher.Num_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
  'frmShowVoucher.dImporte = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
  'frmShowVoucher.sFlg_Status = GridEX1.Value(GridEX1.Columns("Estatus_Letra").Index)
  frmShowVoucher.BUSCAR
  frmShowVoucher.FunctButt1.ChangeProperty "ENABLED", 1, False

  frmShowVoucher.Show vbModal
  Set frmShowVoucher = Nothing

Exit Sub

Resume
errx:
    errores err.Number

End Sub

Private Sub GeneraInfoContable()
On Error GoTo errx
Dim vResp As Variant
Dim sSQL As String

vResp = MsgBox("Confirma Generación Contable de Documento ? ", vbYesNo, "CONFIRMACION")
If vResp = vbNo Then Exit Sub

sSQL = "CN_GENERA_ASIENTO_VENTAS '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"

ExecuteCommandSQL cConnect, sSQL
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
BUSCAR
Exit Sub

errx:
    errores err.Number
End Sub

Private Sub CmdAceptar_Click()
    ReactivarFactura
End Sub

Private Sub cmdCancelar_Click()
    fraReactivarFactura.Visible = False
End Sub

Private Sub txtSer_Docum2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtSer_Docum = StrZero(txtSer_Docum2.Text, 3)
        txtNum_Docum2.SetFocus
    End If
End Sub

Private Sub txtNum_Docum2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNum_Docum2 = StrZero(txtNum_Docum2.Text, 8)
        cmdAceptar.SetFocus
    End If
End Sub


Sub ReactivarFactura()
On Error GoTo fin
Dim sTit As String
Dim sAccion As String, StrSQL As String

    
   sTit = "Reactivar FActura Anulada"
    
   If MsgBox("Confirma Reactivación de factura Anulada...?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
    StrSQL = "EXEC VENTAS_REACTIVA_FACTURA_ANULADA '" & txtCod_TipDoc2.Text & "','" & txtSer_Docum2.Text & "','" & txtNum_Docum2.Text & "'"
        
    ExecuteCommandSQL cConnect, StrSQL
    Me.fraReactivarFactura.Visible = False
    BUSCAR
    
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Sub Factura_sa(Tipo As String, Serie As String)
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String, lvRuta As String, StrSQL As String
Dim rsFactura  As New ADODB.Recordset
     
    Set oo = CreateObject("excel.application")
    
    If Tipo = "FA" Then
        StrSQL = "VENTAS_EMITE_FACTURA_VENTAS_DETA_ROLLO '" & SNum_Corre & "','" & UCase(EnLetras(Trim(CStr(SImp_Total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
        If rsFactura.RecordCount <= 0 Then Exit Sub
        oo.Workbooks.Open vRuta & "\Factura_Tela_Acabada_Rollo_Directa.XLT"
        
    End If
    If Tipo = "BV" Then
        StrSQL = "VENTAS_EMITE_FACTURA_VENTAS_DETA_ROLLO '" & SNum_Corre & "','" & UCase(EnLetras(Trim(CStr(SImp_Total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
        If rsFactura.RecordCount <= 0 Then Exit Sub
        oo.Workbooks.Open vRuta & "\Impresion_Boleta.XLT"
        
    End If

    If Tipo = "ND" Then
       StrSQL = "Ventas_Emite_Documento_Abono_ROLLO '" & SNum_Corre & "','" & UCase(EnLetras(Trim(CStr(SImp_Total)))) & "'"
       Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
       If rsFactura.RecordCount <= 0 Then Exit Sub
       oo.Workbooks.Open vRuta & "\Abono_Textil.XLT"
    End If
    
    If Tipo = "NC" Then
        StrSQL = "Ventas_Emite_Documento_Abono_ROLLO '" & SNum_Corre & "','" & UCase(EnLetras(Trim(CStr(SImp_Total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
        If rsFactura.RecordCount <= 0 Then Exit Sub
        oo.Workbooks.Open vRuta & "\Credito_Textil.XLT"
    End If
          
    oo.DisplayAlerts = False
    oo.Visible = True
    
    oo.run "Reporte", rsFactura, 0, cConnect
    
    Set oo = Nothing
        
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion de La Factura " & err.Description, vbCritical, "Impresion"
End Sub
    
Function IMPRIMIR_DOCSALIDA() As Boolean
Dim StrSQL As String, sNomPartida As String, oPrint As Object
iLin = 0
Set oPrint = New clsPrintFile
    IMPRIMIR_DOCSALIDA = False
    Close #1
    Open "C:\Factura.txt" For Output As #1
    
    Plin Chr(15) & "   "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
       ' Plin "     "

    
    sNomPartida = IMPRIME_CAB_DOCSALIDA
    IMPRIME_DET_DOCSALIDA sNomPartida
    
    Close #1
    oPrint.SendPrint "c:\factura.txt"
    Set oPrint = Nothing
    IMPRIMIR_DOCSALIDA = True
End Function

Function IMPRIME_CAB_DOCSALIDA() As String
Dim RsPro As ADODB.Recordset, RS As ADODB.Recordset
Dim strCadena As String
Dim varDescripcion As String
Dim sOrden As String
Dim iespacios As Long
Dim StrSQL As String

Dim sDirPartida As String, sNomPartida As String, sRucPartida As String, _
    sDirDestino As String, sNomDestino As String, sRucDestino As String
Dim sMarcaPlaca As String, sPeso As String, sMTC As String, sNoLic As String, _
    sNomTransp As String, sempresa As String, _
    SrucEmpresa As String, sventa As String
    
    Plin "     "
    'Plin "     "
    'Plin "     "
    




StrSQL = "EXEC Ventas_Emite_Factura_Ventas_CAB '" & SNum_Corre & "'"
Set RsPro = New ADODB.Recordset
RsPro.ActiveConnection = cConnect
RsPro.CursorLocation = adUseClient
RsPro.CursorType = adOpenStatic
RsPro.Open StrSQL




strCadena = Space(105) & RsPro.Fields("Serie").Value & "-" & RsPro.Fields("Numero").Value
Plin strCadena


strCadena = Space(5) & Format(RsPro.Fields("Fec_EmiDoc").Value, "dd") & Space(6) & Format(RsPro.Fields("Fec_EmiDoc").Value, "mm") & Space(6) & Format(RsPro.Fields("Fec_EmiDoc").Value, "yyyy")
    

Plin strCadena

Plin "     "


 
iespacios = 105 - (22 + Len(sNomDestino))
strCadena = Space(10) & RsPro.Fields("Cliente").Value
Plin strCadena
Plin "     "
Plin "     "
strCadena = Space(10) & RsPro.Fields("Dir_Anexo").Value
Plin strCadena
strCadena = " "
Plin strCadena
Plin "     "
strCadena = Space(10) & RsPro.Fields("Ruc").Value & Space(50) & RsPro.Fields("Des_Venta").Value
Plin strCadena

'Plin "     "
Plin "     "
Plin "     "

   
IMPRIME_CAB_DOCSALIDA = sNomPartida
strCadena = ""
Plin strCadena
Plin strCadena
strCadena = Space(100)
Plin strCadena
strCadena = Space(100)
End Function

Sub IMPRIME_DET_DOCSALIDA(sNomPartida As String)
Dim RS As ADODB.Recordset
Dim varNroReg As Integer
Dim NroReg As Integer
Dim NumTotReg As Integer
Dim strCadena As String
Dim varObserv As String, varObserv1 As String 'Para Observaciones en Guia Manual
Dim i As Integer
Dim iMaxLen As Integer
Dim varDescripcion As String
Dim vFila As Integer
Dim vExcede As Integer
Dim sDescripcion As String, sUnd As String
Dim sDetraccion As String
Dim sletras As String
Dim ssubtotal
Dim sigv
Dim spigv As String
Dim Stotal
Dim smoneda As String
Dim scod_almacen As String
Dim sguia As String
Dim sglosafactura As String
Dim sFdetraccion As String


Set RS = New ADODB.Recordset
RS.ActiveConnection = cConnect
RS.CursorLocation = adUseClient
RS.CursorType = adOpenStatic

    



RS.Open "EXEC Ventas_Emite_Factura_Ventas_Deta '" & SNum_Corre & "','" & UCase(EnLetras(Trim(CStr(GridEX1.Value(GridEX1.Columns("Imp_Total").Index))))) & "'"


sDetraccion = Trim(RS!Glosa)
sletras = Trim(RS!Letras)
ssubtotal = RS!Neto
Stotal = RS!Total
sigv = RS!Igv
spigv = RS!PIgv
smoneda = RS!Moneda
scod_almacen = Trim(RS!Cod_almacen)
sguia = Trim(RS!Guias)
sglosafactura = Trim(RS!GlosaFactura)
sFdetraccion = Trim(RS!Flg_Detraccion)
iMaxLen = 50

If RS.RecordCount Then
    varNroReg = 1
    NroReg = 1
    RS.MoveFirst
    vExcede = 0
    For i = 1 To RS.RecordCount
    sUnd = "KG"
        
        If scod_almacen = "31" Then
            sDescripcion = Trim(RS!Descripcion)
            strCadena = Space(14) & AlExp(RS!KIlos, 18) & Space(3) & AlExp("KG", 3) & Space(7) & AlExp(Mid(sDescripcion, 1, iMaxLen), 50) & Space(5) & AlExp(RS!precio, 10) & Space(6) & AlExp(RS!ImporteTela, 10)
            Plin strCadena
       End If
        If scod_almacen = "30" Then
            strCadena = Space(14) & AlExp(RS!KIlos, 18) & Space(3) & AlExp("KG", 3) & Space(4) & AlExp(Mid(RS!Ruta, 1, 50), 50) & Space(2) & AlExp(RS!precio, 10) & Space(5) & AlExp(RS!ImporteTela, 10)
            Plin strCadena
     End If
        
        If scod_almacen = "30" Then
          sDescripcion = Trim(RS!Descripcion)
            If Len(strCadena) > 0 Then
                vFila = 1
            Do While sDescripcion <> ""
            
                varDescripcion = AlExp(Mid(sDescripcion, 1, iMaxLen), CLng(iMaxLen))
                 
                If vFila = 1 Then
                   Plin Space(42) & varDescripcion  '& Space(15)
                Else
                   Plin Space(42) & varDescripcion
                End If
                sDescripcion = Mid(sDescripcion, iMaxLen + 1)
                NroReg = NroReg + 1
                vFila = vFila + 1
            Loop
        Else
            NroReg = NroReg + 1
            Plin strCadena
        End If
        End If
        If NroReg > 14 Then
            vExcede = 1
            Exit For
        Else
            RS.MoveNext
        End If
    Next
    
    
    If scod_almacen = "31" Then
        NroReg = i
    End If
    

    
    'For i = NroReg To 20
    For i = NroReg To 15
        Plin "     "
    Next
    
    If sglosafactura <> "" Then
            strCadena = Space(30) & "Obs :" & sglosafactura
            Plin strCadena
    Else
        Plin "     "
    End If
    
      If sDetraccion <> "" And scod_almacen = "30" Or sFdetraccion = "S" Then
        '    Plin "     "
            Plin "     "
            Plin "     "
            strCadena = Space(30) & sDetraccion
            Plin strCadena
        Else
            'Plin "     "
             Plin "     "
             Plin "     "
         Plin "     "
        
      End If
      
        'strCadena = Space(10) & sletras
        'Plin strCadena
        
        If scod_almacen = "31" And NroReg = 5 Then
                 Plin "     "
                 Plin "     "
        End If
         If scod_almacen = "31" And NroReg = 6 Then
                 Plin "     "
                 Plin "     "
        End If
        If scod_almacen = "31" And NroReg = 8 Then
                 Plin "     "
                 Plin "     "
                 Plin "     "
        End If
        If scod_almacen = "31" And NroReg = 10 Then
                 Plin "     "
        End If
        If scod_almacen = "31" And NroReg = 11 Then
                 Plin "     "
                 Plin "     "
        End If
         If scod_almacen = "31" And NroReg = 7 Then
                 Plin "     "
                 Plin "     "
                 Plin "     "
                 Plin "     "
        End If
        
        
        If NroReg = 9 Or NroReg = 8 Or NroReg = 11 Or NroReg = 7 Then
             Plin "     "
             Plin "     "
        
        Else
         Plin "     "
         Plin "     "
         Plin "     "
         Plin "     "
         Plin "     "
        
        End If
        
         
         strCadena = ""
         strCadena = Space(12) & sletras
         Plin strCadena
         Plin "     "
         Plin "     "
         Plin "     "
         Plin "     "
         Plin "     "
         strCadena = Space(104) & smoneda & Space(4) & RTrim(AlExp(ssubtotal, 10))
         Plin strCadena
         Plin "     "
         Plin "     "
         strCadena = Space(104) & spigv & "%" & Space(4) & RTrim(AlExp(sigv, 10))
         Plin strCadena
         Plin "     "
         strCadena = Space(104) & smoneda & Space(4) & RTrim(AlExp(Stotal, 10))
         Plin strCadena
         strCadena = Space(18) & sguia
         Plin strCadena


    'IMPRIME_REF_DOCSALIDA sNomPartida
    Plin Chr(12)
    
    If vExcede = 1 Then MsgBox "La cantidad de detalle excede el tamaño de la Guia, algunos datos no se imprimieron, verifique", vbInformation, Me.Caption
    
End If

End Sub




Sub Plin(ByVal Text)
    If IsNull(Text) Then
       Text = ""
    End If
    Print #1, Text
    iLin = iLin + 1
End Sub


Private Function AlExp(Exp As Variant, Longitud As Long) As String
On Error GoTo fin
Dim bEsString As Boolean
    'Alinear Expresion
    bEsString = False
    Select Case VarType(Exp)
    Case vbInteger Or vbLong
        AlExp = Format(Exp, "#,###,##0")
    Case vbDecimal   'Or vbDouble Or vbSingle
        AlExp = Format(Exp, "#,###,##0.00")
    Case vbString
        bEsString = True
        AlExp = Exp
    Case Else
        AlExp = ""
    End Select
    If bEsString Then
        AlExp = Left(AlExp & Space(200), Longitud)
    Else
        If AlExp = "0.00" Then
        AlExp = Right(Space(200) & "", Longitud)
        Else
        AlExp = Right(Space(200) & AlExp, Longitud)
        End If
    End If
Exit Function
fin:
End Function



Private Sub ActualizaFechaRecepcion()
On Error GoTo errores
Dim sSQL As String
Dim sFlg_Pendalidad As String

If txtFecrecepcion.Text = "" Then
    sSQL = "VN_Actualiza_FechaRecepcion '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "',null"
Else
    sSQL = "VN_Actualiza_FechaRecepcion '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & txtFecrecepcion.Text & "'"
End If

ExecuteCommandSQL cConnect, sSQL
GridEX1.Enabled = True
Me.Fra_FechaRecepcion.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub





