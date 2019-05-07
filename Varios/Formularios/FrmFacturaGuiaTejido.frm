VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Begin VB.Form FrmFacturaGuiaTejido 
   Caption         =   "AUTORIZACION DE FACTURACION"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16605
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   16605
   StartUpPosition =   3  'Windows Default
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   8355
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   14737
      _Version        =   262144
      TabCount        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabMinWidth     =   5505
      Tabs            =   "FrmFacturaGuiaTejido.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   7965
         Left            =   30
         TabIndex        =   28
         Top             =   360
         Width           =   16515
         _ExtentX        =   29131
         _ExtentY        =   14049
         _Version        =   262144
         TabGuid         =   "FrmFacturaGuiaTejido.frx":00D2
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&IMPRIMIR"
            Height          =   495
            Left            =   15240
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   50
            Width           =   1185
         End
         Begin GridEX20.GridEX GridEX3 
            Height          =   7290
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   16305
            _ExtentX        =   28760
            _ExtentY        =   12859
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
            HeaderFontName  =   "MS Sans Serif"
            HeaderFontBold  =   0   'False
            HeaderFontSize  =   8.25
            HeaderFontWeight=   400
            FontName        =   "MS Sans Serif"
            FontBold        =   0   'False
            FontSize        =   8.25
            FontWeight      =   400
            ColumnHeaderHeight=   285
            FormatStylesCount=   7
            FormatStyle(1)  =   "FrmFacturaGuiaTejido.frx":00FA
            FormatStyle(2)  =   "FrmFacturaGuiaTejido.frx":0232
            FormatStyle(3)  =   "FrmFacturaGuiaTejido.frx":02E2
            FormatStyle(4)  =   "FrmFacturaGuiaTejido.frx":0396
            FormatStyle(5)  =   "FrmFacturaGuiaTejido.frx":046E
            FormatStyle(6)  =   "FrmFacturaGuiaTejido.frx":0526
            FormatStyle(7)  =   "FrmFacturaGuiaTejido.frx":0606
            ImageCount      =   0
            PrinterProperties=   "FrmFacturaGuiaTejido.frx":0626
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   7965
         Left            =   30
         TabIndex        =   8
         Top             =   360
         Width           =   16515
         _ExtentX        =   29131
         _ExtentY        =   14049
         _Version        =   262144
         TabGuid         =   "FrmFacturaGuiaTejido.frx":07FE
         Begin GridEX20.GridEX GridEX4 
            Height          =   2055
            Left            =   840
            TabIndex        =   34
            Top             =   4920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   3625
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ScrollToolTipColumn=   ""
            HeaderFontSize  =   8.25
            FontSize        =   8.25
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "FrmFacturaGuiaTejido.frx":0826
            Column(2)       =   "FrmFacturaGuiaTejido.frx":08EE
            FormatStylesCount=   6
            FormatStyle(1)  =   "FrmFacturaGuiaTejido.frx":0992
            FormatStyle(2)  =   "FrmFacturaGuiaTejido.frx":0ACA
            FormatStyle(3)  =   "FrmFacturaGuiaTejido.frx":0B7A
            FormatStyle(4)  =   "FrmFacturaGuiaTejido.frx":0C2E
            FormatStyle(5)  =   "FrmFacturaGuiaTejido.frx":0D06
            FormatStyle(6)  =   "FrmFacturaGuiaTejido.frx":0DBE
            ImageCount      =   0
            PrinterProperties=   "FrmFacturaGuiaTejido.frx":0E9E
         End
         Begin GridEX20.GridEX GridEX6 
            Height          =   2055
            Left            =   3600
            TabIndex        =   35
            Top             =   4920
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   3625
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ScrollToolTipColumn=   ""
            HeaderFontSize  =   8.25
            FontSize        =   8.25
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "FrmFacturaGuiaTejido.frx":1076
            Column(2)       =   "FrmFacturaGuiaTejido.frx":113E
            FormatStylesCount=   6
            FormatStyle(1)  =   "FrmFacturaGuiaTejido.frx":11E2
            FormatStyle(2)  =   "FrmFacturaGuiaTejido.frx":131A
            FormatStyle(3)  =   "FrmFacturaGuiaTejido.frx":13CA
            FormatStyle(4)  =   "FrmFacturaGuiaTejido.frx":147E
            FormatStyle(5)  =   "FrmFacturaGuiaTejido.frx":1556
            FormatStyle(6)  =   "FrmFacturaGuiaTejido.frx":160E
            ImageCount      =   0
            PrinterProperties=   "FrmFacturaGuiaTejido.frx":16EE
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&GENERAR FACTURAS"
            Height          =   495
            Left            =   15240
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   50
            Width           =   1185
         End
         Begin GridEX20.GridEX GridEX2 
            Height          =   7050
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   12435
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
            HeaderFontName  =   "MS Sans Serif"
            HeaderFontBold  =   0   'False
            HeaderFontSize  =   8.25
            HeaderFontWeight=   400
            FontName        =   "MS Sans Serif"
            FontBold        =   0   'False
            FontSize        =   8.25
            FontWeight      =   400
            ColumnHeaderHeight=   285
            ColumnsCount    =   2
            Column(1)       =   "FrmFacturaGuiaTejido.frx":18C6
            Column(2)       =   "FrmFacturaGuiaTejido.frx":198E
            FormatStylesCount=   8
            FormatStyle(1)  =   "FrmFacturaGuiaTejido.frx":1A32
            FormatStyle(2)  =   "FrmFacturaGuiaTejido.frx":1B6A
            FormatStyle(3)  =   "FrmFacturaGuiaTejido.frx":1C1A
            FormatStyle(4)  =   "FrmFacturaGuiaTejido.frx":1CCE
            FormatStyle(5)  =   "FrmFacturaGuiaTejido.frx":1DA6
            FormatStyle(6)  =   "FrmFacturaGuiaTejido.frx":1E5E
            FormatStyle(7)  =   "FrmFacturaGuiaTejido.frx":1F3E
            FormatStyle(8)  =   "FrmFacturaGuiaTejido.frx":1FEA
            ImageCount      =   0
            PrinterProperties=   "FrmFacturaGuiaTejido.frx":209A
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   45
            Left            =   0
            TabIndex        =   24
            Top             =   600
            Width           =   16455
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   7965
         Left            =   30
         TabIndex        =   7
         Top             =   360
         Width           =   16515
         _ExtentX        =   29131
         _ExtentY        =   14049
         _Version        =   262144
         TabGuid         =   "FrmFacturaGuiaTejido.frx":2272
         Begin VB.CheckBox chkExpandir 
            Caption         =   "EXPANDIR"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   15240
            TabIndex        =   27
            Top             =   120
            Width           =   1155
         End
         Begin VB.TextBox txtGuia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   18
            Top             =   90
            Width           =   1635
         End
         Begin VB.TextBox txtSerieFac 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10710
            TabIndex        =   17
            Top             =   60
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "EXPANDIR"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   7680
            TabIndex        =   16
            Top             =   120
            Width           =   1155
         End
         Begin VB.TextBox txtNroFactura 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11460
            TabIndex        =   15
            Top             =   60
            Width           =   1875
         End
         Begin VB.TextBox txtSerie 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   780
            MaxLength       =   3
            TabIndex        =   14
            Top             =   90
            Width           =   585
         End
         Begin VB.CommandButton CmdAsignar 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   8850
            TabIndex        =   11
            Top             =   1740
            Width           =   555
         End
         Begin VB.CommandButton cmdLiberaGuia 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   8850
            TabIndex        =   10
            Top             =   2340
            Width           =   555
         End
         Begin GridEX20.GridEX GridEX5 
            Height          =   7260
            Left            =   9420
            TabIndex        =   12
            Top             =   600
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   12806
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
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
            HeaderFontName  =   "MS Sans Serif"
            HeaderFontBold  =   0   'False
            HeaderFontSize  =   8.25
            HeaderFontWeight=   400
            FontName        =   "MS Sans Serif"
            FontBold        =   0   'False
            FontSize        =   8.25
            FontWeight      =   400
            ColumnHeaderHeight=   285
            FormatStylesCount=   7
            FormatStyle(1)  =   "FrmFacturaGuiaTejido.frx":229A
            FormatStyle(2)  =   "FrmFacturaGuiaTejido.frx":23D2
            FormatStyle(3)  =   "FrmFacturaGuiaTejido.frx":2482
            FormatStyle(4)  =   "FrmFacturaGuiaTejido.frx":2536
            FormatStyle(5)  =   "FrmFacturaGuiaTejido.frx":260E
            FormatStyle(6)  =   "FrmFacturaGuiaTejido.frx":26C6
            FormatStyle(7)  =   "FrmFacturaGuiaTejido.frx":27A6
            ImageCount      =   0
            PrinterProperties=   "FrmFacturaGuiaTejido.frx":27C6
         End
         Begin GridEX20.GridEX GridEX1 
            Height          =   7290
            Left            =   0
            TabIndex        =   13
            Top             =   600
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   12859
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
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
            HeaderFontName  =   "MS Sans Serif"
            HeaderFontBold  =   0   'False
            HeaderFontSize  =   8.25
            HeaderFontWeight=   400
            FontName        =   "MS Sans Serif"
            FontBold        =   0   'False
            FontSize        =   8.25
            FontWeight      =   400
            ColumnHeaderHeight=   285
            FormatStylesCount=   7
            FormatStyle(1)  =   "FrmFacturaGuiaTejido.frx":299E
            FormatStyle(2)  =   "FrmFacturaGuiaTejido.frx":2AD6
            FormatStyle(3)  =   "FrmFacturaGuiaTejido.frx":2B86
            FormatStyle(4)  =   "FrmFacturaGuiaTejido.frx":2C3A
            FormatStyle(5)  =   "FrmFacturaGuiaTejido.frx":2D12
            FormatStyle(6)  =   "FrmFacturaGuiaTejido.frx":2DCA
            FormatStyle(7)  =   "FrmFacturaGuiaTejido.frx":2EAA
            ImageCount      =   0
            PrinterProperties=   "FrmFacturaGuiaTejido.frx":2ECA
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   45
            Left            =   0
            TabIndex        =   23
            Top             =   480
            Width           =   16455
         End
         Begin VB.Label Label4 
            Caption         =   "NRO GUIA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1410
            TabIndex        =   22
            Top             =   120
            Width           =   825
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9090
            TabIndex        =   21
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label5 
            Caption         =   "ASIGNA FACTURA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9270
            TabIndex        =   20
            Top             =   150
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "SERIE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   19
            Top             =   150
            Width           =   825
         End
      End
   End
   Begin VB.Frame FraBuscar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16560
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   11580
         TabIndex        =   32
         Top             =   255
         Width           =   3315
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   10530
         TabIndex        =   31
         Top             =   255
         Width           =   1050
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&BUSCAR"
         Height          =   495
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1185
      End
      Begin VB.ComboBox Cbo_Almacen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   5910
         TabIndex        =   2
         Top             =   195
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Format          =   73007105
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   7680
         TabIndex        =   3
         Top             =   195
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         Format          =   73007105
         CurrentDate     =   37543
      End
      Begin VB.Label Label9 
         Caption         =   "CLIENTE"
         Height          =   255
         Left            =   9840
         TabIndex        =   33
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "RANGO DE FECHA EMISION"
         Height          =   360
         Left            =   3810
         TabIndex        =   5
         Top             =   225
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "ALMACEN"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmFacturaGuiaTejido"
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
Dim StrSQL As String
Public CODIGO As String
Public Descripcion As String
Public TipoAdd As String
Dim sCod_TipoFact  As String

Dim sSer_Factura_Orig As String
Dim sNum_Factura_Orig As String

Private Sub Check1_Click()
    If GridEX1.RowCount = 0 Then Exit Sub
    With GridEX1
        Select Case CBool(Check1.Value)
            Case True: .ExpandAll
            Case False: .CollapseAll
        End Select
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
End Sub

Private Sub chkExpandir_Click()
    If GridEX5.RowCount = 0 Then Exit Sub
    With GridEX5
        Select Case CBool(chkExpandir.Value)
            Case True: .ExpandAll
            Case False: .CollapseAll
        End Select
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
End Sub


Private Sub CmdAsignar_Click()

Dim sSQL As String
On Error GoTo Error_Handler
Dim oGroup As GridEX20.JSGroup

    Call Asigna_Numero_Factura
    Call BuscarPendientes
    'Call BuscarGuiasAutorizadas
  
Exit Sub

Resume

Error_Handler:
errores err.Number

End Sub

Private Sub cmdBuscar_Click()
    Call BuscarPendientes
    Call BuscarGuiasConfacturas
   'Call BuscarGuiasAutorizadas
End Sub

Private Sub cmdGuardar_Click()

End Sub

Private Sub cmdLiberaGuia_Click()

Dim sSQL As String
On Error GoTo Error_Handler
Dim oGroup As GridEX20.JSGroup

    Call libera_Guia_de_Factura
    Call BuscarPendientes
    Call BuscarGuiasConfacturas
  
Exit Sub

Resume

Error_Handler:

  errores err.Number
   
  'If ColIndex = GridEX5.Columns("Sel").Index Then
  '   GridEX1.Value(GridEX5.Columns("sel").Index) = 0
  'End If
End Sub
Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TX_Cliente  ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txtAbr_Cliente.Text = CODIGO
        txtNom_Cliente.Text = Descripcion
        StrSQL = "SELECT Cod_Cliente_Tex As Cod_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        txtAbr_Cliente.Tag = DevuelveCampo(StrSQL, cConnect)

        SendKeys "{TAB}"
        CODIGO = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub



Private Sub Form_Load()

  dtpFecEmiIni.Value = Date - 10
  dtpFecEmiFin.Value = Date
 
  FillAlmacen
  
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
 '''' 2do tab
 Set GridEX4.ADORecordset = CargarRecordSetDesconectado("select Cod_CondVent,Des_CondVent as Descripcion from lg_condvent", cConnect)
    
  GridEX4.ColumnAutoResize = True
  
  GridEX4.ActAsDropDown = True
  GridEX4.BoundColumnIndex = 1
  GridEX4.ReplaceColumnIndex = 2
   
  
  GridEX4.Columns("Cod_CondVent").Visible = False
  
  Set GridEX6.ADORecordset = CargarRecordSetDesconectado("select Cod_Moneda as cod_Moneda,Nom_Moneda as Descripcion from tg_moneda", cConnect)
    
  GridEX6.ColumnAutoResize = True

  GridEX6.ActAsDropDown = True
  GridEX6.BoundColumnIndex = 1
  GridEX6.ReplaceColumnIndex = 2
  
  GridEX6.Columns("Cod_Moneda").Visible = False



End Sub

Private Sub BuscarPendientes()

On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

   
sSQL = "EXEC CN_MUESTRA_GUIA_SS_TEJIDO_PENDIENTES_ASIGNAR_FACTURAS '" & Left(Cbo_Almacen, 2) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & Trim(txtAbr_Cliente.Tag) & "','" & Trim(txtSerie.Text) & "','" & Trim(txtGuia.Text) & "'"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
  
Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("cliente").Index, jgexSortAscending)

SetColoresizquierda

GridEX1.DefaultGroupMode = jgexDGMCollapsed

'If dtpFecEmiIni.Value <> "" Then
'    GridEX1.DefaultGroupMode = jgexDGMExpanded
'End If

If GridEX1.RowCount < 35 Then
    GridEX1.DefaultGroupMode = jgexDGMExpanded
End If


If GridEX1.RowCount > 0 Then
    GridEX1.Row = 1
End If

GridEX1.ContinuousScroll = True

Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub
Private Sub BuscarGuiasConfacturas()
On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSQL = "EXEC CN_MUESTRA_GUIA_SS_TEJIDO_FACTURA_ASIGNADA '" & Left(Cbo_Almacen, 2) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & Trim(txtAbr_Cliente.Tag) & "','',''"

GridEX5.ClearFields

GridEX5.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX5.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
  
Set oGroup = GridEX5.Groups.Add(GridEX5.Columns("cliente").Index, jgexSortAscending)
'GridEX5.Columns("Fecha_Guia").Width = 900
'GridEX5.Columns("Fecha_Guia").EditType = jgexEditNone
SetColoresDerecha
'
GridEX5.DefaultGroupMode = jgexDGMCollapsed

If GridEX5.RowCount < 35 Then
    GridEX5.DefaultGroupMode = jgexDGMExpanded
End If

If GridEX5.RowCount > 0 Then
    GridEX5.Row = 1
End If

GridEX5.ContinuousScroll = True

Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub

Private Sub Form_Resize()
'GridEX1.Width = Me.Width - 300
GridEX1.Height = Me.Height - 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Sub AfterColEdit_Prendas(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler
Dim oGroup As GridEX20.JSGroup
Select Case ColIndex
  Case Is = GridEX1.Columns("Num_Factura").Index
    Call Asigna_Numero_Factura
    Call BuscarPendientes
    Call BuscarGuiasConfacturas
   
  End Select
  
  
Exit Sub

Resume

Error_Handler:

  errores err.Number
End Sub
Private Sub Asigna_Numero_Factura()
Dim sSQL As String
Dim num_factura As String
Dim Serie As String
On Error GoTo errx

    Serie = Trim(txtSerieFac.Text)
    num_factura = Trim(txtNroFactura.Text)

    If Trim(Serie) = "" Or Trim(num_factura) = "" Then
    Call MsgBox("Debe Ingresar numero Factura", vbInformation, "Mensaje")
    Exit Sub
    End If



      sSQL = "TX_RELACIONA_GUIA_FACTURA_SS_TEJIDO '$','$','$','$','$'"
      
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("movimiento").Index), _
                       Serie, _
                       num_factura, _
                       GridEX1.Value(GridEX1.Columns("cod_cliente").Index))


    ExecuteCommandSQL cConnect, sSQL
    
    
    
Exit Sub
errx:
    errores err.Number
End Sub

Private Sub libera_Guia_de_Factura()
Dim sSQL As String
Dim num_factura As String
Dim Serie As String
On Error GoTo errx
    
    Serie = ""
    num_factura = ""
    
      sSQL = "TX_LIBERA_GUIA_DE_FACTURA_SS_TEJIDO '$','$','$','$','$'"
      
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX5.Value(GridEX5.Columns("movimiento").Index), _
                       Serie, _
                       num_factura, _
                       GridEX5.Value(GridEX5.Columns("cod_cliente").Index))
    ExecuteCommandSQL cConnect, sSQL
    
    
Exit Sub
errx:
    errores err.Number
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)

Dim strGroupCaption As String

'If RowBuffer.RowType = jgexRowTypeGroupHeader Then
'    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
'    RowBuffer.GroupCaption = strGroupCaption
'End If

End Sub


Private Sub SetColoresizquierda()

Dim fmtCon As JSFmtCondition
Dim fmtCond2 As JSFmtCondition
Dim fmtCond3 As JSFmtCondition

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("GUIA").Index, jgexNotEqual, "")
    
    With GridEX1.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Guias"
            Set fmtCon = .GroupCondition
    End With
    fmtCon.SetCondition GridEX1.Columns("cliente").Index, jgexNotEqual, ""
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
End Sub
Private Sub SetColoresDerecha()

Dim fmtCon As JSFmtCondition
Dim fmtCond2 As JSFmtCondition
Dim fmtCond3 As JSFmtCondition

Set fmtCon = GridEX5.FmtConditions.Add(GridEX5.Columns("GUIA").Index, jgexNotEqual, "")
    
    With GridEX5.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Guias"
            Set fmtCon = .GroupCondition
    End With
    fmtCon.SetCondition GridEX5.Columns("cliente").Index, jgexNotEqual, ""
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
End Sub

Private Sub FillAlmacen()

Dim rstAux As ADODB.Recordset
Dim StrSQL As String


StrSQL = "EXEC sa_MUESTRA_ALMACENES_FACTURACION '" & vusu & "'"


Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
Cbo_Almacen.Clear
With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        Cbo_Almacen.AddItem !Cod_almacen & " " & !nom_almacen
        .MoveNext
    Loop
    .Close
End With
If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
Set rstAux = Nothing

End Sub

Private Sub txtAbr_Cliente_Change()
        txtAbr_Cliente.Tag = ""
    
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            StrSQL = "SELECT Nom_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(StrSQL, cConnect)
            StrSQL = "SELECT Cod_Cliente_Tex As Cod_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtAbr_Cliente.Tag = DevuelveCampo(StrSQL, cConnect)
            SendKeys "{TAB}"
            
        End If
    End If
End Sub

Public Function CargaValores(ByRef ObjTemp As Object) As Boolean
    ObjTemp.txtAbr_Cliente.Text = txtAbr_Cliente.Text
    ObjTemp.txtAbr_Cliente.Tag = txtAbr_Cliente.Tag
    ObjTemp.txtDEs_cliente.Text = txtNom_Cliente.Text
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

Private Sub txtEstGrafico_Change()
Call BuscarPendientes
End Sub

Private Sub txtGuia_Change()
Call BuscarPendientes
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If
End Sub

Public Sub BUSCA_CLIENTE(tipo As Integer)
    Select Case tipo
    
        Case 1:
                    StrSQL = "EXEC SA_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                    If Trim(txtNom_Cliente.Text) <> "" Then
                    
                     Call BuscarPendientes
                     Call BuscarGuiasConfacturas
                    End If
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "EXEC SA_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC SA_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.Cargar_Datos
                    'oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
'                         OptCliPend.SetFocus
                         CODIGO = "": Descripcion = ""
                         BuscarPendientes
                         BuscarGuiasConfacturas
                         
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
                    

            StrSQL = "SELECT Cod_Cliente_Tex As Cod_Cliente FROM TX_CLIENTE WHERE  Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtAbr_Cliente.Tag = DevuelveCampo(StrSQL, cConnect)
                                
    End Select
    
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    Else
        Call SoloNumeros(txtNroFactura, KeyAscii, False, 0, 8)
    End If
End Sub

Private Sub txtNroFactura_LostFocus()
    txtNroFactura.Text = Format(Trim(txtNroFactura.Text), "00000000")
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtGuia.SetFocus
    Else
        Call SoloNumeros(txtSerie, KeyAscii, False, 0, 3)
    End If
    
End Sub

Private Sub txtSerie_LostFocus()
    txtSerie.Text = Format(Trim(txtSerie.Text), "000")
End Sub

Private Sub txtSerieFac_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtNroFactura.SetFocus
    Else
        Call SoloNumeros(txtSerieFac, KeyAscii, False, 0, 3)
    End If
    
End Sub
Private Sub txtSerieFac_LostFocus()
    txtSerieFac.Text = Format(Trim(txtSerieFac.Text), "000")
End Sub

''''************************************************************************************************************
''''Termina primer tab 1
''''************************************************************************************************************

'Private Sub Form_Resize()
'    GridEX2.Width = Me.Width - 300
'End Sub
Private Sub DtFecVencimiento_Change()
  GridEX2.ClearFields
'  dtpFecEmiIni.Value = ""
'  dtpFecEmiFin.Value = ""
End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX2.ClearFields
'  If Trim(dtpFecEmiIni.Value) <> "" Then
'    'dtpFecEmiFin.Value = dtpFecEmiIni
'  End If
End Sub

'Private Sub BuscarGuiasAutorizadas()
'
'On Error GoTo drdepurar
'
'Dim sSQL As String
'Dim oGroup As GridEX40.JSGroup
'Dim oFormat As JSFormatStyle
'
'  sSQL = "CN_MUESTRA_GUIA_SS_TEJIDO_AUTORIZACION_FACTURACION '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIn & "','" & dtpFecEmiFin & "','" & Trim(txtAbr_Cliente.Tag) & "','',''"
'
'GridEX2.ClearFields
'
'GridEX2.DefaultGroupMode = jgexDGMExpanded
'bCargaGRid = False
'Set GridEX2.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
'
'Set oGroup = GridEX2.Groups.Add(GridEX2.Columns("Fac_Cli").Index, jgexSortAscending)
'
'MuestraSubTotales
'GridEX2.BackColorRowGroup = &H80000005
'
'GridEX2.ColumnHeaderHeight = 500
'
'
'GridEX2.Columns("fecha").Width = 975
'GridEX2.Columns("fecha").Visible = False
'
'GridEX2.Columns("CLIENTE").Visible = False
'GridEX2.Columns("Orden").Width = 1015
'
'
'GridEX2.Columns("nro_guia").Visible = True
'
''GridEX2.Columns("observaciones").Width = 2000
'GridEX2.Columns("nro_Guia").Width = 1240
'GridEX2.Columns("nro_Guia").Caption = "GUIA"
'GridEX2.Columns("cod_item").Width = 825
'GridEX2.Columns("cod_estcli").Width = 1125
'GridEX2.Columns("cod_estcli").Caption = "ESTILO"
'GridEX2.Columns("grafico").Width = 825
'GridEX2.Columns("grafico").Caption = "GRAFICO"
'
'GridEX2.Columns("SERVICIO").Width = 1000
'GridEX2.Columns("SERVICIO").Caption = "SERVICIO"
'
'GridEX2.Columns("LOTE").Width = 825
'GridEX2.Columns("LOTE").Caption = "LOTE"
'GridEX2.Columns("observaciones").Visible = False
'GridEX2.Columns("num_movstk").Visible = False
'GridEX2.Columns("cod_item").Visible = False
'GridEX2.Columns("item").Width = 6000
'GridEX2.Columns("moneda").Width = 900
'
'GridEX2.Columns("pu").Width = 840
'GridEX2.Columns("CANTIDAD").Caption = "Prendas"
'GridEX2.Columns("CANTIDAD").Width = 765
'
'GridEX2.Columns("monto despacho").Width = 855
'GridEX2.Columns("SEL").Width = 450
'
'GridEX2.Columns("Fac_Cli").Width = 0
'
'GridEX2.Columns("Ser_Factura").Width = 500
'GridEX2.Columns("Num_Factura").Width = 900
'
'
'GridEX2.Columns("COD_CONDVENT").Visible = False
'GridEX2.Columns("Cod_Moneda").Visible = False
''GridEX2.Columns("Num_movstk").Visible = False
'GridEX2.Columns("SER_ORDCOMP").Visible = False
'GridEX2.Columns("SEC_ORDCOMP").Visible = False
'GridEX2.Columns("COD_ORDCOMP").Visible = False
'
'GridEX2.Columns("Ser_Factura").Caption = "Serie"
'GridEX2.Columns("Num_Factura").Caption = "Nro Factura"
'
'
'GridEX2.Columns("PU").Format = "#######0.0000"
'GridEX2.Columns("PU").Caption = "Precio"
'
'GridEX2.Columns("importe").Format = "#######0.00"
'GridEX2.Columns("importe").Caption = "Monto Desp"
'
'
'GridEX2.Columns("SEL").ColumnType = jgexCheckBox
'GridEX2.Columns("SEL").Visible = True
'GridEX2.Columns("SEL").EditType = jgexEditCheckBox
'GridEX2.Columns("SEL").Width = 500
'
'With GridEX2.Columns("Condicion_Venta")
'  .TextAlignment = jgexAlignLeft
'  .EditType = jgexEditCombo
'  Set .DropDownControl = GridEX4
'End With
'
'With GridEX2.Columns("moneda")
'  .TextAlignment = jgexAlignLeft
'  .EditType = jgexEditCombo
'  Set .DropDownControl = GridEX6
'End With
'
''With grxDatos
''    Set oGroup01 = .Groups.Add(.Columns("DESCRIPCION").Index, jgexSortAscending)
''
''    If chkExpandir.Value = Checked Then
''        .DefaultGroupMode = jgexDGMExpanded
''    Else
''        .DefaultGroupMode = jgexDGMCollapsed
''    End If
''
''    .GroupFooterStyle = jgexTotalsGroupFooter
''
''End With
''
'SetColores
'
'GridEX2.DefaultGroupMode = jgexDGMCollapsed
'
'If dtpFecEmiIni.Value <> "" Then
'    GridEX2.DefaultGroupMode = jgexDGMExpanded
'End If
'
'GridEX2.ContinuousScroll = True
'
'Exit Sub
'Resume
'drdepurar:
'  errores Err.Number
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
'End Sub
'
'
'
'Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'    Dim Msg As Variant
'    Select Case ActionName
'    Case "BUSCAR"
'      BUSCAR
'    Case "AUTORIZARPAGO"
'        If GridEX2.RowCount = 0 Then Exit Sub
'        Msg = MsgBox("¿Esta seguro de autorizar pago?", vbYesNo)
'        If Msg = vbNo Then Exit Sub
'        AUTORIZAR
'    Case "SALIR"
'       Unload Me
'    End Select
'End Sub
'
'Private Sub GridEX2_AfterColEdit(ByVal ColIndex As Integer)
'
'  If Left(Cbo_Almacen, 2) = "20" Then
'    AfterColEdit_Prendas (ColIndex)
'
'  End If
'
'End Sub
'Sub AfterColEdit_Prendas(ByVal ColIndex As Integer)
'
'Dim sSQL As String
'On Error GoTo Error_Handler
'
'Dim oGroup As GridEX40.JSGroup
'Dim cadena As String
'cadena = ""
'Select Case ColIndex
'
'  Case Is = GridEX2.Columns("Sel").Index
'
'      sSQL = "SA_VENTAS_CAMBIO_ESTADO_DOCALM_SERVICIOS '$','$','$','$','$',$,'$',$,$,'$','$','$',$,'$','$'"
'
'      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
'                       GridEX2.Value(GridEX2.Columns("num_movstk").Index), _
'                       GridEX2.Value(GridEX2.Columns("Ser_Factura").Index), _
'                       GridEX2.Value(GridEX2.Columns("Num_Factura").Index), _
'                       GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index), _
'                       GridEX2.Value(GridEX2.Columns("Pre_Unitario").Index), _
'                       GridEX2.Value(GridEX2.Columns("Cod_Moneda").Index), _
'                       0, _
'                       0, _
'                       GridEX2.Value(GridEX2.Columns("Ser_ordcomp").Index), _
'                       GridEX2.Value(GridEX2.Columns("Cod_ordcomp").Index), _
'                       GridEX2.Value(GridEX2.Columns("Sec_OrdComp").Index), _
'                       GridEX2.Value(GridEX2.Columns("num_prendas").Index), _
'                       cadena, _
'                       cadena)
'
'    ExecuteCommandSQL cConnect, sSQL
'
'  Case Is = GridEX2.Columns("Pre_Unitario").Index
'  'EXEC SA_MAN_PRECIOS_ORDEN_SERVICIO '00001','002','043315','095','13','1.4'
'  sSQL = "SA_MAN_PRECIOS_ORDEN_SERVICIO '$','$','$','$','$',$ " '·,'$',$,$,'$','$','$',$,'$','$'"
'        sSQL = VBsprintf(sSQL, GridEX2.Value(GridEX2.Columns("COD_CLIENTE_SA").Index), _
'                       GridEX2.Value(GridEX2.Columns("Ser_ordcomp").Index), _
'                       GridEX2.Value(GridEX2.Columns("Cod_ordcomp").Index), _
'                       GridEX2.Value(GridEX2.Columns("Sec_OrdComp").Index), _
'                       GridEX2.Value(GridEX2.Columns("COD_PROCESO_SA").Index), _
'                       GridEX2.Value(GridEX2.Columns("Pre_Unitario").Index))
'
'
'    GridEX2.Value(GridEX2.Columns("Monto Despacho").Index) = GridEX2.Value(GridEX2.Columns("Pre_Unitario").Index) * GridEX2.Value(GridEX2.Columns("num_prendas").Index)
'    GridEX2.Value(GridEX2.Columns("sel").Index) = 0
'
'    ExecuteCommandSQL cConnect, sSQL
'
'  Case Is = GridEX2.Columns("num_prendas").Index
'    GridEX2.Value(GridEX2.Columns("Monto Despacho").Index) = GridEX2.Value(GridEX2.Columns("Pre_Unitario").Index) * GridEX2.Value(GridEX2.Columns("num_prendas").Index)
'    GridEX2.Value(GridEX2.Columns("sel").Index) = 0
'  Case Is = GridEX2.Columns("Ser_Factura").Index
'    GridEX2.Value(GridEX2.Columns("Fac_Cli").Index) = GridEX2.Value(GridEX2.Columns("Ser_Factura").Index) & "-" & GridEX2.Value(GridEX2.Columns("Num_Factura").Index) & "  " & GridEX2.Value(GridEX2.Columns("Nom_Cliente").Index)
'    GridEX2.Groups.Clear
'    Set oGroup = GridEX2.Groups.Add(GridEX2.Columns("Fac_Cli").Index, jgexSortAscending)
'    GridEX2.Value(GridEX2.Columns("sel").Index) = 0
'  Case Is = GridEX2.Columns("Num_Factura").Index
'    If Trim(GridEX2.Value(GridEX2.Columns("Ser_Factura").Index)) = "" Then GridEX2.Value(GridEX2.Columns("Ser_Factura").Index) = "001"
'    GridEX2.Value(GridEX2.Columns("Fac_Cli").Index) = GridEX2.Value(GridEX2.Columns("Ser_Factura").Index) & "-" & GridEX2.Value(GridEX2.Columns("Num_Factura").Index) & "  " & GridEX2.Value(GridEX2.Columns("Nom_Cliente").Index)
'    GridEX2.Groups.Clear
'    Set oGroup = GridEX2.Groups.Add(GridEX2.Columns("Fac_Cli").Index, jgexSortAscending)
'    GridEX2.Value(GridEX2.Columns("sel").Index) = 0
'
'  End Select
'Exit Sub
'
'Resume
'
'Error_Handler:
'
'  errores Err.Number
'
'  If ColIndex = GridEX2.Columns("Sel").Index Then
'     GridEX2.Value(GridEX2.Columns("sel").Index) = 0
'  End If
'End Sub
'
'
'Private Sub GridEX2_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX40.JSRetBoolean)
'
'If Left(Cbo_Almacen, 2) = "20" Then
'
'  Select Case ColIndex
'
'    Case Is = GridEX2.Columns("Ser_Factura").Index
'      If Trim(GridEX2.Value(GridEX2.Columns("Ser_Factura").Index)) = "" Then GridEX2.Value(GridEX2.Columns("Ser_Factura").Index) = ""
'      Cancel = False
'    Case Is = GridEX2.Columns("Num_Factura").Index
'      If Trim(GridEX2.Value(GridEX2.Columns("Num_Factura").Index)) = "" Then GridEX2.Value(GridEX2.Columns("Num_Factura").Index) = ""
'      Cancel = False
'    Case Is = GridEX2.Columns("SEL").Index
'      Cancel = False
'    Case Is = GridEX2.Columns("Pre_Unitario").Index
'      Cancel = False
'    Case Is = GridEX2.Columns("Condicion_Venta").Index
'      Cancel = False
'    Case Is = GridEX2.Columns("Moneda").Index
'      Cancel = False
'   Case Else
'      Cancel = True
'   End Select
'
'End If
'
'End Sub
'
'Private Sub GridEX2_Click()
'
''On Error Resume Next
'    Dim ColIndex As Long
'    Dim oRowData As JSRowData
'    Dim SGRUPO As String
'    Dim iRow As Long
'    Dim i As Long
'    Dim sCaptionGroup As String
'
'    bCargaGRid = True
'
'        If GridEX2.RowCount > 0 Then
'        ColIndex = GridEX2.Col
'
'        If Not GridEX2.IsGroupItem(GridEX2.Row) Then
'            If UCase(GridEX2.Columns(ColIndex).Key) = "SEL" Then
'                bClickColSelec = True
'                SendKeys "{ENTER}"
'            End If
'        Else
'            If GridEX2.IsGroupItem(GridEX2.Row) Then
'            End If
'        End If
'    End If
'End Sub
'
'Private Sub GridEX2_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
'    Dim ocol As JSColumn
'    Dim oRow As JSRowData
'    Dim vCurrentRow As Variant
'    Dim oRowGroup As JSRowData
'    Dim sProveedor As String
'
'    iColAnterior = LastCol
'    iRowAnterior = LastRow
'
'    If GridEX2.Row <> 0 Then
'        Set oRow = GridEX2.GetRowData(GridEX2.Row)
'    End If
'
'    If GridEX2.RowCount > 0 Then
'      On Error Resume Next
'
'     lbldesItem.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("item").Index)), "", GridEX2.Value(GridEX2.Columns("item").Index))
''      lbComb.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("Comb").Index)), "", GridEX2.Value(GridEX2.Columns("Comb").Index))
''      lbCalidad.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("Calidad").Index)), "", GridEX2.Value(GridEX2.Columns("Calidad").Index))
''      lbRollos.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("Numero_Rollos").Index)), "", GridEX2.Value(GridEX2.Columns("Numero_Rollos").Index))
''      If lbCod_Color.Visible Then lbDes_Color.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("Color").Index)), "", GridEX2.Value(GridEX2.Columns("Color").Index))
'      lbGuia.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("nro_Guia").Index)), "", GridEX2.Value(GridEX2.Columns("nro_Guia").Index))
'      lbobservacion.Caption = IIf(IsNull(GridEX2.Value(GridEX2.Columns("Observaciones").Index)), "", GridEX2.Value(GridEX2.Columns("Observaciones").Index))
''
'    End If
'End Sub
'
'Private Sub GridEX2_RowFormat(RowBuffer As GridEX40.JSRowData)
'
'Dim strGroupCaption As String
'
'If RowBuffer.RowType = jgexRowTypeGroupHeader Then
'    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
'    RowBuffer.GroupCaption = strGroupCaption
'End If
'
'If GridEX2.RowCount = 0 Then Exit Sub
'Dim fmtConDIA_Programado As JSFmtCondition
'Set fmtConDIA_Programado = GridEX2.FmtConditions.Add(GridEX2.Columns("Pre_Unitario").Index, jgexEqual, 0#, 0)
'With fmtConDIA_Programado.FormatStyle
'    .ForeColor = &H8000&
'    .FontSize = 8
'    .BackColor = &H80000018 'vbYellow
'End With
'
'
''Private Sub GridEX2_RowFormat(RowBuffer As GridEX40.JSRowData)
''End Sub
'
'End Sub
'
'Private Sub MuestraSubTotales()
'Dim colTemp As JSColumn
'
'GridEX2.GroupFooterStyle = jgexTotalsGroupFooter
'Set colTemp = GridEX2.Columns("Moneda")
'colTemp.AggregateFunction = jgexAggregateNone
'colTemp.TotalRowPrefix = "SUB TOTAL "
'
'GridEX2.GroupFooterStyle = jgexTotalsGroupFooter
'Set colTemp = GridEX2.Columns("num_prendas")
'colTemp.AggregateFunction = jgexSum
'colTemp.TotalRowPrefix = ""
'
'GridEX2.GroupFooterStyle = jgexTotalsGroupFooter
'Set colTemp = GridEX2.Columns("Monto Despacho")
'colTemp.AggregateFunction = jgexSum
'colTemp.TotalRowPrefix = ""
'
'End Sub
'
'Private Sub SetColores()
'
'Dim fmtCon As JSFmtCondition
'Dim fmtCond2 As JSFmtCondition
'Dim fmtCond3 As JSFmtCondition
'
'Set fmtCon = GridEX2.FmtConditions.Add(GridEX2.Columns("SEL").Index, jgexEqual, -1)
'
'    With GridEX2.FmtConditions
'            .ApplyGroupCondition = True
'            .ShowGroupConditionCount = True
'            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
'            Set fmtCon = .GroupCondition
'    End With
'    fmtCon.SetCondition GridEX2.Columns("SEL").Index, jgexEqual, -1
'    fmtCon.FormatStyle.FontBold = True
'    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
'
'End Sub
'
'
'Private Sub AUTORIZAR()
''On Error GoTo errorx
'On Error GoTo SALTO_ERROR
'Dim sSQL As String
'Dim aMess(4), i As Integer
'
'If Left(Cbo_Almacen, 2) = "20" Then
'    Call GuardaCambios
'End If
'
'If Left(Cbo_Almacen, 2) = "20" Then
'  ExecuteCommandSQL cConnect, "SA_VENTAS_GENERA_DOCUM_AUTORIZADOS_SERVICIOS '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
'End If
'Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
'Call BUSCAR
'
''Exit Sub
''Resume
''errorx:
''    ErrorHandler Err, "Autoriza Documentos"
'Exit Sub
'SALTO_ERROR:
'MsgBox Err.Description, vbCritical, Me.Caption
'
'End Sub
'
'Private Sub GuardaCambios()
'
'  On Error GoTo SALTO_ERROR
'  Dim sSQL As String
'  Dim rs As ADODB.Recordset
'
'  If GridEX2.RowCount > 0 Then
'        GridEX2.Update
'        Set rs = GridEX2.ADORecordset
'
'        rs.MoveFirst
'
'        Do While Not rs.EOF
'
'           If rs.Fields("sel").Value <> 0 Then
'
'                     sSQL = "SA_VENTAS_CAMBIO_ESTADO_DOCALM_SERVICIOS '$','$','$','$','$',$,'$',$,$,'$','$','$',$,'$','$'"
'                     sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
'                                       Trim(rs.Fields("num_movstk").Value), _
'                                       Trim(rs.Fields("Ser_Factura").Value), _
'                                       Trim(rs.Fields("Num_Factura").Value), _
'                                       Trim(rs.Fields("Cod_CondVent").Value), _
'                                       Trim(rs.Fields("Pre_Unitario").Value), _
'                                       Trim(rs.Fields("Cod_Moneda").Value), _
'                                       0, _
'                                       0, _
'                                       Trim(rs.Fields("Ser_ordcomp").Value), _
'                                       Trim(rs.Fields("Cod_ordcomp").Value), _
'                                       Trim(rs.Fields("Sec_OrdComp").Value), _
'                                       Trim(rs.Fields("num_prendas").Value), _
'                                       "", _
'                                       "")
'                    ExecuteCommandSQL cConnect, sSQL
'           End If
'
'           rs.MoveNext
'
'        Loop
'
'        rs.MoveFirst
'        Set GridEX2.ADORecordset = rs
'        GridEX2.SetFocus
'    End If
'    'Call MsgBox("Registro seleccionados se insertaron con exito", vbOKOnly, "Mensaje")
'    Exit Sub
'SALTO_ERROR:
'    MsgBox Err.Description, vbCritical, Me.Caption
'
'End Sub
'
''Sub Cambio_Nro_Factura()
''
''Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
''
''  GridEX2.Redraw = False
''
''  lvSW = True
''
''  Doc = GridEX2.Value(GridEX2.Columns("Cod_Doc").Index)
''  Serie = GridEX2.Value(GridEX2.Columns("Ser_Docum").Index)
''  Nro_Factura = GridEX2.Value(GridEX2.Columns("Num_Docum_Ventas").Index)
''
''  GridEX2.MoveFirst
''  For i = 0 To GridEX2.RowCount
''    If Doc = GridEX2.Value(GridEX2.Columns("Cod_Doc").Index) Then
''      If lvSW Then iPos = GridEX2.Row
''      lvSW = False
''      GridEX2.Value(GridEX2.Columns("Ser_Docum").Index) = Serie
''      GridEX2.Value(GridEX2.Columns("Nro_Docum_Ventas").Index) = Nro_Factura
''    End If
''    GridEX2.MoveNext
''  Next i
''
''  GridEX2.Row = iPos
''
''  GridEX2.Redraw = True
''
''  SendKeys "{TAB}"
''
''End Sub
'
'
'Private Sub GridEX4_Click()
'
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'
'  GridEX2.Redraw = False
'
'  lvSw = True
'
'  serie = GridEX2.Value(GridEX2.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX2.Value(GridEX2.Columns("Num_Factura").Index)
'
'
'  GridEX2.MoveFirst
'  For i = 0 To GridEX2.RowCount
'    If serie = GridEX2.Value(GridEX2.Columns("Ser_Factura").Index) And Nro_Factura = GridEX2.Value(GridEX2.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX2.Row
'      lvSw = False
'      GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index) = GridEX4.Value(GridEX4.Columns("Cod_CondVent").Index)
'      GridEX2.Value(GridEX2.Columns("Condicion_Venta").Index) = GridEX4.Value(GridEX4.Columns("Descripcion").Index)
'    End If
'    GridEX2.MoveNext
'  Next i
'
'  GridEX2.Row = iPos
'
'  GridEX2.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub
'
'Private Sub GridEX6_Click()
'
'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSw As Boolean
'
'  GridEX2.Redraw = False
'
'  serie = GridEX2.Value(GridEX2.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX2.Value(GridEX2.Columns("Num_Factura").Index)
'  lvSw = True
'  GridEX2.MoveFirst
'  For i = 0 To GridEX2.RowCount
'    If serie = GridEX2.Value(GridEX2.Columns("Ser_Factura").Index) And Nro_Factura = GridEX2.Value(GridEX2.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX2.Row
'      lvSw = False
'      GridEX2.Value(GridEX2.Columns("Cod_Moneda").Index) = GridEX6.Value(GridEX6.Columns("Cod_Moneda").Index)
'      GridEX2.Value(GridEX2.Columns("Moneda").Index) = GridEX6.Value(GridEX6.Columns("Descripcion").Index)
'    End If
'    GridEX2.MoveNext
'  Next i
'
'  GridEX2.Row = iPos
'
'  GridEX2.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub
'
'
''Private Sub FillAlmacen()
''
''Dim rstAux As ADODB.Recordset
''Dim strSQL As String
''
''strSQL = "SA_VENTAS_AYUDA_ALMACENES_PRENDAS"
''
''Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
''Cbo_Almacen.Clear
''With rstAux
''    If .RecordCount > 0 Then .MoveFirst
''    Do Until .EOF
''        Cbo_Almacen.AddItem !Cod_almacen & " " & !nom_almacen
''        .MoveNext
''    Loop
''    .Close
''End With
''If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
''Set rstAux = Nothing
''
''End Sub
'''''*******************************************************************************************
'''''***********termina segundo tab
'''''*******************************************************************************************
'
'
'
'
