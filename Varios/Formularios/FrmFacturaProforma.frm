VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFacturaProforma 
   Caption         =   "Factura Proforma"
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraBuscar 
      BackColor       =   &H00C0FFFF&
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
      Height          =   765
      Left            =   60
      TabIndex        =   39
      Top             =   0
      Width           =   13920
      Begin VB.TextBox txtCod_OrdComp 
         Height          =   285
         Left            =   6300
         TabIndex        =   43
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtSer_OrdComp 
         Height          =   285
         Left            =   5760
         TabIndex        =   42
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   1605
         TabIndex        =   41
         Top             =   240
         Width           =   3360
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   900
         TabIndex        =   40
         Top             =   240
         Width           =   690
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   525
         Left            =   7800
         TabIndex        =   44
         Top             =   120
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   926
         Custom          =   $"FrmFacturaProforma.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   500
         ControlSeparator=   40
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "N° O/S"
         Height          =   255
         Left            =   5160
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cliente:"
         Height          =   210
         Left            =   240
         TabIndex        =   45
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos Del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   18
      Top             =   840
      Width           =   13935
      Begin VB.ComboBox cmbCiudad_Cli 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2040
         Width           =   4035
      End
      Begin VB.TextBox Txt_Direccion_Cli 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   23
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox Txt_Atencion_Cli 
         Height          =   285
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   22
         Top             =   960
         Width           =   4095
      End
      Begin VB.ComboBox cmbViaTransporte_Cli 
         Height          =   315
         ItemData        =   "FrmFacturaProforma.frx":012C
         Left            =   7560
         List            =   "FrmFacturaProforma.frx":012E
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   3555
      End
      Begin VB.TextBox Txt_Destino_Cli 
         Height          =   645
         Left            =   7560
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Txt_NroFP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Para:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Atención:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "País"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lbl_Pais_Cli 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label lbl_para_Cli 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Via Transporte"
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Destino"
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbl_Fecha 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7560
         TabIndex        =   28
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "%Incremento"
         Height          =   255
         Left            =   6240
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblPorcInc 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7560
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "N° FP"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox Txt_Flete 
      Height          =   375
      Left            =   3780
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox Txt_Seguro 
      Height          =   375
      Left            =   6660
      TabIndex        =   13
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   8280
      Width           =   14055
      Begin VB.CommandButton cmd_busdatobanco 
         Caption         =   "Datos Banco"
         Height          =   375
         Left            =   10440
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmd_registra 
         Caption         =   "Nuevo Banco"
         Height          =   375
         Left            =   12240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl_beneficiario 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label26 
         Caption         =   "Beneficiario"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_swift 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7800
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Swift"
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_direccion_datobanco 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7800
         TabIndex        =   7
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lbl_cuenta 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_banco 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label23 
         Caption         =   "Banco"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   60
      TabIndex        =   15
      Top             =   3360
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Telas"
      TabPicture(0)   =   "FrmFacturaProforma.frx":0130
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GridEX2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Colores"
      TabPicture(1)   =   "FrmFacturaProforma.frx":014C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridEX1"
      Tab(1).ControlCount=   1
      Begin GridEX20.GridEX GridEX1 
         Height          =   3780
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   6668
         Version         =   "2.0"
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
         Column(1)       =   "FrmFacturaProforma.frx":0168
         Column(2)       =   "FrmFacturaProforma.frx":0230
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmFacturaProforma.frx":02D4
         FormatStyle(2)  =   "FrmFacturaProforma.frx":040C
         FormatStyle(3)  =   "FrmFacturaProforma.frx":04BC
         FormatStyle(4)  =   "FrmFacturaProforma.frx":0570
         FormatStyle(5)  =   "FrmFacturaProforma.frx":0648
         FormatStyle(6)  =   "FrmFacturaProforma.frx":0700
         FormatStyle(7)  =   "FrmFacturaProforma.frx":07E0
         FormatStyle(8)  =   "FrmFacturaProforma.frx":088C
         ImageCount      =   0
         PrinterProperties=   "FrmFacturaProforma.frx":093C
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   3780
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   13680
         _ExtentX        =   24130
         _ExtentY        =   6668
         Version         =   "2.0"
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
         Column(1)       =   "FrmFacturaProforma.frx":0B14
         Column(2)       =   "FrmFacturaProforma.frx":0BDC
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmFacturaProforma.frx":0C80
         FormatStyle(2)  =   "FrmFacturaProforma.frx":0DB8
         FormatStyle(3)  =   "FrmFacturaProforma.frx":0E68
         FormatStyle(4)  =   "FrmFacturaProforma.frx":0F1C
         FormatStyle(5)  =   "FrmFacturaProforma.frx":0FF4
         FormatStyle(6)  =   "FrmFacturaProforma.frx":10AC
         FormatStyle(7)  =   "FrmFacturaProforma.frx":118C
         FormatStyle(8)  =   "FrmFacturaProforma.frx":1238
         ImageCount      =   0
         PrinterProperties=   "FrmFacturaProforma.frx":12E8
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6375
      Top             =   4905
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label17 
      Caption         =   "Flete"
      Height          =   255
      Left            =   2820
      TabIndex        =   54
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "Seguro"
      Height          =   375
      Left            =   5700
      TabIndex        =   53
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "CIF US$"
      Height          =   255
      Left            =   8340
      TabIndex        =   52
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label lbl_CIF 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9540
      TabIndex        =   51
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label LBL_ETIQ 
      Caption         =   "FOB"
      Height          =   255
      Left            =   300
      TabIndex        =   50
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label lbl_FOB 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1140
      TabIndex        =   49
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label lbl_IdproformaFactura 
      Height          =   255
      Left            =   11820
      TabIndex        =   48
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_cod_banco 
      Caption         =   "Label22"
      Height          =   375
      Left            =   14580
      TabIndex        =   47
      Top             =   6480
      Width           =   1095
   End
End
Attribute VB_Name = "FrmFacturaProforma"
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
Public CODIGO As String
Public descripcion As String
Public sNroCuenta As String
Public campo3 As String
Public campo4 As String
Public campo5 As String
Public campo6 As String
Public sidcliente As String
Public TipoAdd As String
Dim sCod_TipoFact  As String
Dim sSer_Factura_Orig As String
Dim sNum_Factura_Orig As String
Public rstAux1 As ADODB.Recordset
Public rstAux2 As ADODB.Recordset
Dim rsFactura  As New ADODB.Recordset
Dim rsx_Cabecera_Imp As ADODB.Recordset
Dim rsx_Detalle_Imp As ADODB.Recordset


'lbl_banco--------Descripcion
'lbl_cuenta-------sNroCuenta
'lbl_idcliente----sidcliente
'lbl_cod_banco ---Codigo


Private Sub Cmd_Buscar_Click()
'BuscarTela
'BUSCAR_PROFORMAFACTURA (Txt_NroFP.Text)
'FillCiudad
'FillViaTransporte
End Sub
'Private Sub BuscarCabecera()
'    On Error GoTo SAlTO_ERROR
'
'    Dim oRs As New Recordset
'
'
'
'    strSQL = "EXEC Ups_Muestra_Proforma '" & sIdProforma & "'"
'    Set oRs = CargarRecordSetDesconectado(strSQL, cConnect)
'    If oRs.RecordCount > 0 Then
'
'
'
'    lbl_idOT.Caption = oRs.Fields("IdOrdenTrabajoKey")
'    lbl_IdProforma.Caption = oRs.Fields("IdProformaKey")
'    DtpIni.Value = oRs.Fields("FechaInicioTra")
'
'
'
'End Sub
Private Sub FillCiudad()

    strSQL = "Select Cod_Ubigeo,NombreCiudad From cn_PaisCiudad Where Cod_Pais='" & Left(lbl_Pais_Cli.Caption, 3) & "'"
    
    Set rstAux2 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cmbCiudad_Cli.Clear
    With rstAux2
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cmbCiudad_Cli.AddItem !Cod_Ubigeo & Space(5) & !NombreCiudad
        .MoveNext
    Loop

    End With

    BuscaCombo1 "C", 5, cmbCiudad_Cli
End Sub
Private Sub FillViaTransporte()

    strSQL = "SELECT idViaTransporteKey,NombreVia FROM Tx_MViaTransporte"
    
    Set rstAux1 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cmbViaTransporte_Cli.Clear
    With rstAux1
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cmbViaTransporte_Cli.AddItem !idViaTransporteKey & Space(5) & !NombreVia
        .MoveNext
    Loop

    End With

    BuscaCombo1 "C", 5, cmbViaTransporte_Cli
End Sub

Private Sub cmd_busdatobanco_Click()
'If Trim(txtAbr_Cliente.Text) = "" Then
'    MsgBox "Debe elegir un cliente", vbCritical, "Error"
'    txtAbr_Cliente.SetFocus
'Else
    Call BUSCA_DATOS_BANCO
'End If
End Sub

Private Sub cmd_registra_Click()
frmDatoCtaCliente.Show 1
End Sub

Private Sub Form_Resize()
    GridEX1.Width = Me.Width - 300
End Sub
'Private Sub DtFecVencimiento_Change()
'  GridEX1.ClearFields
'  dtpFecEmiIni.Value = ""
'  dtpFecEmiFin.Value = ""
'End Sub

Private Sub CmdAceptar_Click()
    GuardarDatos
End Sub

Private Sub GuardarDatos()
End Sub


Private Sub cmdLugEnt_Click()
End Sub

'Private Sub Command1_Click()
'    Me.fraDatosAdicionales.Visible = False
'End Sub

'Private Sub dtpFecEmiIni_Change()
'  GridEX1.ClearFields
'  If Trim(dtpFecEmiIni.Value) <> "" Then
'    dtpFecEmiFin.Value = dtpFecEmiIni
'  End If
'End Sub

Private Sub Form_Load()

  
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  
  

End Sub
Private Sub BUSCAR_PROFORMAFACTURA(ByVal sIdProforma As String)
    On Error GoTo SAlTO_ERROR
    
    Dim oRs As New Recordset
   
       
    
    strSQL = "EXEC Usp_Busca_ProformaFactura '" & sIdProforma & "'"
    Set oRs = CargarRecordSetDesconectado(strSQL, cConnect)
    If oRs.RecordCount > 0 Then
    

        lbl_para_Cli.Caption = oRs.Fields("Nom_Cliente")
        lbl_Pais_Cli.Caption = oRs.Fields("Pais_Base")
        If oRs.Fields("Pais") <> "" Then
            lbl_Pais_Cli.Caption = oRs.Fields("Pais")
        End If
        Txt_Direccion_Cli.Text = oRs.Fields("Des_LugEntr")
        FillCiudad
        FillViaTransporte
        
        lbl_Fecha.Caption = oRs.Fields("FechaOperativa")
        lblPorcInc.Caption = oRs.Fields("PorcIncremento")
        lbl_FOB.Caption = oRs.Fields("Fob")
        lbl_CIF.Caption = oRs.Fields("Cif")
        Txt_Flete.Text = oRs.Fields("Flete")
        Txt_Seguro.Text = oRs.Fields("Seguro")
        lbl_IdproformaFactura.Caption = sIdProforma
        
        If Trim(oRs.Fields("Ciudad")) = "" Then
        cmbCiudad_Cli.ListIndex = -1
        Else
        'cmbCiudad_Cli.Text = Trim(oRs.Fields("Ciudad"))
        BuscaCombo1 Trim(oRs.Fields("Ciudad")), 1, cmbCiudad_Cli
        End If
        
        If Trim(oRs.Fields("Via")) = "" Then
            cmbViaTransporte_Cli.ListIndex = -1
        Else
            'cmbViaTransporte_Cli.Text = Trim(oRs.Fields("Via"))
            BuscaCombo1 Trim(oRs.Fields("Via")), 1, cmbViaTransporte_Cli
        End If
        
        Txt_Destino_Cli.Text = Trim(oRs.Fields("DestinoCliente"))
        Txt_Atencion_Cli.Text = Trim(oRs.Fields("AtencionCliente"))
        If Trim(oRs.Fields("DireccionCliente")) <> "" Then
           Txt_Direccion_Cli.Text = Trim(oRs.Fields("DireccionCliente"))
        End If
        
        'DATOS DEL BANCO
        lbl_banco.Caption = Trim(oRs.Fields("Nom_Banco"))
        lbl_cod_banco.Caption = Trim(oRs.Fields("cod_banco"))
        lbl_cuenta.Caption = Trim(oRs.Fields("nro_cuenta"))
        lbl_direccion_datobanco.Caption = Trim(oRs.Fields("Direccion"))
        lbl_direccion_datobanco.Caption = Trim(oRs.Fields("Direccion"))
        lbl_beneficiario.Caption = Trim(oRs.Fields("beneficiario"))
        lbl_swift.Caption = Trim(oRs.Fields("cod_swift"))
        
               
    Else
    
        lbl_para_Cli.Caption = ""
        Txt_Atencion_Cli.Text = ""
        Txt_Direccion_Cli.Text = ""
        lbl_Pais_Cli.Caption = ""
        cmbCiudad_Cli.ListIndex = -1
        lbl_Fecha.Caption = ""
        cmbViaTransporte_Cli.ListIndex = -1
        Txt_Destino_Cli.Text = ""
        lbl_IdproformaFactura.Caption = ""
        cmbCiudad_Cli.ListIndex = -1
        cmbViaTransporte_Cli.ListIndex = -1
        Txt_Destino_Cli.Text = ""
        Txt_Atencion_Cli.Text = ""
        Txt_Direccion_Cli.Text = ""
        lbl_banco.Caption = ""
        lbl_cod_banco.Caption = ""
        lbl_cuenta.Caption = ""
        lbl_direccion_datobanco.Caption = ""
        lbl_direccion_datobanco.Caption = ""
        lbl_beneficiario.Caption = ""
        lbl_swift.Caption = ""
        

    End If
    
    Exit Sub
SAlTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption


End Sub
Private Sub BuscarTela()
On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle


sSQL = "EXEC Usp_Busca_Detalle_ProformaFactura_Tela '" & Txt_NroFP.Text & "'"



Set GridEX2.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
  

GridEX2.Columns("Des_Tela").Caption = "Descripcion"
GridEX2.Columns("Des_Tela").Width = 4500
GridEX2.Columns("Des_Tela_x_Pais").Caption = "Descripcion_Pais"
GridEX2.Columns("Des_Tela_x_Pais").Width = 4500
GridEX2.Columns("Can_Pedida").Caption = "Cantidad"
GridEX2.Columns("Cod_Tela").Visible = False


Exit Sub
Resume
drDepurar:
  errores Err.Number
End Sub
Private Sub BuscarColor()

On Error GoTo drDepurar
Dim sCod_Tela As String

If GridEX2.RowCount <> 0 Then

        sCod_Tela = GridEX2.Value(GridEX2.Columns("Cod_Tela").Index)
        
        If sCod_Tela <> "" Then
        
                Dim sSQL As String
                Dim oGroup As GridEX20.JSGroup
                Dim oFormat As JSFormatStyle
                
                
                
                sSQL = "Usp_Busca_Detalle_ProformaFactura '" & Txt_NroFP.Text & "','" & sCod_Tela & "'"
                
                
                'GridEX1.ClearFields
                
                'GridEX1.DefaultGroupMode = jgexDGMExpanded
                'bCargaGRid = False
                Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
                  
                
        '        GridEX1.Columns("Des_Tela").Caption = "Descripción"
        '        GridEX1.Columns("Des_Tela").Width = 3300
                GridEX1.Columns("Des_Tela_x_Pais").Caption = "Descripción Pais"
                GridEX1.Columns("Des_Tela_x_Pais").Width = 3300
                GridEX1.Columns("Des_Color").Caption = "Color"
                GridEX1.Columns("Des_Color").Width = 3300
                
                GridEX1.Columns("Des_Color_x_Pais").Caption = "Color Pais"
                GridEX1.Columns("Des_Color_x_Pais").Width = 3300
                
                GridEX1.Columns("PartidaArancelaria").Caption = "Partida Arancelaria"
                GridEX1.Columns("Pre_Unitario").Caption = "Precio Kg"
                GridEX1.Columns("Can_Pedida").Caption = "Cantidad"
                
                GridEX1.Columns("Des_Tela").Visible = False
                GridEX1.Columns("Des_Tela_x_Pais").Visible = False
                GridEX1.Columns("Cod_Tela").Visible = False
                GridEX1.Columns("Cod_Color").Visible = False
                
                
                
                
                'GridEX1.DefaultGroupMode = jgexDGMCollapsed
                
                
                'GridEX1.DefaultGroupMode = jgexDGMExpanded
                
                
                'GridEX1.ContinuousScroll = True
        
        End If
End If
Exit Sub
Resume
drDepurar:
  errores Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
        BuscaIdProformaFactura
        BuscarTela
        BUSCAR_PROFORMAFACTURA (Txt_NroFP.Text)

    Case "GRABARFACTURAPROFORMA"
        If MsgBox("Desea Grabar los cambios para este registro?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
                Grabar
        End If
        
   Case "IMPRIMIR"
    On Error GoTo xerror:
         Txt_NroFP = Format(Txt_NroFP, "00000000")
         Set rsx_Cabecera_Imp = New ADODB.Recordset
         Set rsx_Detalle_Imp = New ADODB.Recordset
         Set rsx_Cabecera_Imp = CargarRecordSetDesconectado("Exec Carga_Cabecera_FActura_Proforma '" & Txt_NroFP.Text & "'", cConnect)
         Set rsx_Detalle_Imp = CargarRecordSetDesconectado("Exec Carga_Detalle_FActura_Proforma '" & Txt_NroFP.Text & "'", cConnect)
         Dim oo As Object
        Screen.MousePointer = 11
        Set oo = CreateObject("Excel.Application")
        oo.Workbooks.Open vRuta & "\Factura_Proforma_joc.xlt"
        oo.Visible = True
        oo.DisplayAlerts = False
        oo.Run "Cabecera_Factura_Proforma", rsx_Cabecera_Imp, rsx_Detalle_Imp, Txt_NroFP, cConnect
        Set oo = Nothing
        Screen.MousePointer = 0
        Exit Sub
xerror:
             Screen.MousePointer = 0
            ErrorHandler Err, "Factura Proforma"
            Set oo = Nothing
            Exit Sub

        
        
    Case "SALIR"
       Unload Me
    End Select
End Sub
Private Sub BuscaIdProformaFactura()
    On Error GoTo SAlTO_ERROR
    
    Dim oRsCli As New Recordset
   
        
    
    strSQL = "EXEC Usp_Busca_ProfTrabajo '" & Trim(txtAbr_Cliente) & "','" & Trim(txtSer_OrdComp.Text) & "','" & Trim(txtCod_OrdComp.Text) & "'"
    Set oRsCli = CargarRecordSetDesconectado(strSQL, cConnect)
    If oRsCli.RecordCount > 0 Then
    

        Txt_NroFP.Text = oRsCli.Fields("IdFacturaProforma")
        
               
    Else
    
        
        Txt_NroFP.Text = ""
        

    End If
    
    Exit Sub
SAlTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
Private Sub Grabar()
Dim i As Integer
On Error GoTo hand

    strSQL = "EXEC USP_GRABAR_CAB_FACTURAPROFORMA '" & lbl_IdproformaFactura.Caption & "','" & Trim(Txt_Atencion_Cli.Text) & "','" & _
        Trim(Txt_Direccion_Cli.Text) & "','" & _
        Left(cmbCiudad_Cli.Text, 6) & "','" & _
        Left(cmbViaTransporte_Cli.Text, 2) & "','" & _
        Txt_Destino_Cli.Text & "'," & _
        Txt_Flete.Text & "," & _
        Txt_Seguro.Text & "," & _
        lbl_CIF.Caption & "," & _
        lbl_cuenta.Caption & ",'" & _
        lbl_cod_banco.Caption & "'"
        
            
        Call ExecuteSQL(cConnect, strSQL)
    
  
Exit Sub
hand:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub
Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)

    On Error GoTo Error_Handler

    Dim oGroup As GridEX20.JSGroup
    Select Case ColIndex
      Case Is = GridEX1.Columns("Des_Color_x_Pais").Index
            Edita_Nombre_Color
            BuscarColor
      Case Is = GridEX1.Columns("PartidaArancelaria").Index
            Edita_PartidaArancelaria
            BuscarColor
      End Select
    Exit Sub
    Resume

Error_Handler:
      errores Err.Number

End Sub
Sub Edita_PartidaArancelaria()
   Dim sSQL As String, sPartidaArancelaria As String
   
   sPartidaArancelaria = IIf(IsNull(GridEX1.Value(GridEX1.Columns("PartidaArancelaria").Index)) = True, "", GridEX1.Value(GridEX1.Columns("PartidaArancelaria").Index))
   
   If Trim(sPartidaArancelaria) <> "" Then

    sSQL = "Usp_Edita_PartidaArancelaria '$' , '$' , '$' , '$', '$' "
    sSQL = VBsprintf(sSQL, Left(Trim(lbl_Pais_Cli.Caption), 4), GridEX2.Value(GridEX2.Columns("Cod_Tela").Index), GridEX1.Value(GridEX1.Columns("Cod_Color").Index), _
                          Trim(GridEX1.Value(GridEX1.Columns("PartidaArancelaria").Index)), Trim(Txt_NroFP.Text))

        ExecuteCommandSQL cConnect, sSQL
  End If
End Sub
Sub Edita_Nombre_Color()
   Dim sSQL As String, sDesColorxPais As String
   
   sDesColorxPais = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Des_Color_x_Pais").Index)) = True, "", GridEX1.Value(GridEX1.Columns("Des_Color_x_Pais").Index))
   
   If Trim(sDesColorxPais) <> "" Then

    sSQL = "Usp_Edita_NombreColorPais '$' , '$' , '$' , '$' , '$' "
    sSQL = VBsprintf(sSQL, Left(Trim(lbl_Pais_Cli.Caption), 4), GridEX2.Value(GridEX2.Columns("Cod_Tela").Index), GridEX1.Value(GridEX1.Columns("Cod_Color").Index), _
                          Trim(GridEX1.Value(GridEX1.Columns("Des_Color_x_Pais").Index)), Trim(Txt_NroFP.Text))

        ExecuteCommandSQL cConnect, sSQL
  End If
End Sub

Sub Edita_Nombre_Tela()
   Dim sSQL As String, sDesTelaxPais As String
   
   sDesTelaxPais = IIf(IsNull(GridEX2.Value(GridEX2.Columns("Des_Tela_x_Pais").Index)) = True, "", GridEX2.Value(GridEX2.Columns("Des_Tela_x_Pais").Index))
   
   If Trim(sDesTelaxPais) <> "" Then

    sSQL = "Usp_Edita_NombrePais '$' , '$' , '$', '$' "
    sSQL = VBsprintf(sSQL, Left(Trim(lbl_Pais_Cli.Caption), 4), GridEX2.Value(GridEX2.Columns("Cod_Tela").Index), _
                          Trim(GridEX2.Value(GridEX2.Columns("Des_Tela_x_Pais").Index)), Trim(Txt_NroFP.Text))

    ExecuteCommandSQL cConnect, sSQL
    
  End If

End Sub

Private Sub CargarDatos()
End Sub



Private Sub DatosAdic_Click()

'Dim serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
'
'  GridEX1.Redraw = False
'
'  lvSW = True
'
'  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'
'
'  GridEX1.MoveFirst
'  For i = 0 To GridEX1.RowCount
'    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSW Then iPos = GridEX1.Row
'      lvSW = False
'        GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
'        GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
'        GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
'        GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
'        GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
'        GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
'        GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
'        GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
'        GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
'        GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
'        GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
'        GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cConnect), vbLong)
'        GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
'        GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
'    End If
'    GridEX1.MoveNext
'  Next i
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
    
  
End Sub


Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

    Select Case ColIndex
        Case Is = GridEX1.Columns("Can_Pedida").Index
            Cancel = True
        Case Is = GridEX1.Columns("Des_Tela").Index
           Cancel = True
        Case Is = GridEX1.Columns("Des_Color").Index
           Cancel = True
        Case Is = GridEX1.Columns("Total").Index
           Cancel = True
        Case Is = GridEX1.Columns("Pre_Unitario").Index
           Cancel = True
        Case Else
           Cancel = False
     End Select
     

End Sub

Private Sub GridEX1_Click()

'    Dim ColIndex As Long
'    Dim oRowData As JSRowData
'    Dim SGRUPO As String
'    Dim iRow As Long
'    Dim i As Long
'    Dim sCaptionGroup As String
'
'    bCargaGRid = True
'
'        If GridEX1.RowCount > 0 Then
'        ColIndex = GridEX1.Col
'
'        If Not GridEX1.IsGroupItem(GridEX1.Row) Then
'            If UCase(GridEX1.Columns(ColIndex).Key) = "SEL" Then
'                bClickColSelec = True
'                SendKeys "{ENTER}"
'            End If
'
'        Else
'            If GridEX1.IsGroupItem(GridEX1.Row) Then
'            End If
'        End If
'    End If
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
'    Dim ocol As JSColumn
'    Dim oRow As JSRowData
'    Dim vCurrentRow As Variant
'    Dim oRowGroup As JSRowData
'    Dim sProveedor As String
'
'    iColAnterior = LastCol
'    iRowAnterior = LastRow
'
'    If GridEX1.Row <> 0 Then
'        Set oRow = GridEX1.GetRowData(GridEX1.Row)
'    End If
'
'    If GridEX1.RowCount > 0 Then
'      On Error Resume Next
'      lbDesTela.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Tela").Index)), "", GridEX1.Value(GridEX1.Columns("Tela").Index))
'      lbComb.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Comb").Index)), "", GridEX1.Value(GridEX1.Columns("Comb").Index))
'      lbCalidad.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Calidad").Index)), "", GridEX1.Value(GridEX1.Columns("Calidad").Index))
'      lbRollos.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index)), "", GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index))
'      If lbCod_Color.Visible Then lbDes_Color.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Color").Index)), "", GridEX1.Value(GridEX1.Columns("Color").Index))
'      lbGuia.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("nro_Guia").Index)), "", GridEX1.Value(GridEX1.Columns("nro_Guia").Index))
'      lbObservacion.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Observaciones").Index)), "", GridEX1.Value(GridEX1.Columns("Observaciones").Index))
'    End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)

'Dim strGroupCaption As String
'
'If RowBuffer.RowType = jgexRowTypeGroupHeader Then
'    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
'    RowBuffer.GroupCaption = strGroupCaption
'End If

End Sub

Private Sub MuestraSubTotales()

End Sub

Private Sub SetColores()

End Sub


Private Sub Autorizar()

End Sub



Private Sub GridEX2_AfterColEdit(ByVal ColIndex As Integer)
    On Error GoTo Error_Handler
    
    Dim oGroup As GridEX20.JSGroup
    
    Select Case ColIndex
      Case Is = GridEX2.Columns("Des_Tela_x_Pais").Index
            Edita_Nombre_Tela
            BuscarTela
            BuscarColor
            
      End Select
    Exit Sub
    Resume
    
Error_Handler:
      errores Err.Number
End Sub

Private Sub GridEX2_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Select Case ColIndex
        Case Is = GridEX2.Columns("Des_Tela").Index
           Cancel = True
        Case Is = GridEX2.Columns("Can_Pedida").Index
           Cancel = True

        Case Else
           Cancel = False
     End Select
End Sub

Public Sub BuscaCondVent(Opcion As String)
End Sub


Public Sub BuscaLugEnt(Opcion As String)
End Sub


Public Sub BuscaModoTransporte(Opcion As String)
End Sub

Public Sub BuscaTerminoVent(Opcion As String)
End Sub

Private Sub txtCod_Vendor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
    Call BuscarColor
End If
End Sub

Private Sub txt_buscadatosbanco_Change()

End Sub

'Private Sub txt_buscadatosbanco_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Trim(txt_buscadatosbanco.Text) = "" Then
'            Call BUSCA_DATOS_BANCO(3)
'        Else
'            Call BUSCA_DATOS_BANCO(2)
'        End If
'    End If
'End Sub
Public Sub BUSCA_DATOS_BANCO()
Dim strSQL As String

        
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    
                    oTipo.SQuery = "EXEC USP_BUSCA_DATO_BANCO '" & Trim(txtAbr_Cliente.Text) & "'"
   
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         lbl_cod_banco.Caption = Trim(CODIGO)
                         lbl_banco.Caption = Trim(descripcion)
                         lbl_cuenta.Caption = campo3
                         lbl_direccion_datobanco.Caption = campo4
                         lbl_swift.Caption = campo5
                         lbl_beneficiario.Caption = campo6
                         'lbl_banco--------Descripcion
                        'lbl_cuenta-------sNroCuenta
                        'lbl_idcliente----sidcliente
                        'lbl_cod_banco ---Codigo

                         CODIGO = "": descripcion = "": campo3 = "": campo4 = "": campo5 = "": campo6 = ""
                         
                    End If
                    
                    Set oTipo = Nothing
                    Set rs = Nothing
    
    
End Sub

Private Sub Txt_Flete_Change()
Dim dFlete As Double, dSeguro As Double

dFlete = IIf(IsNumeric(Txt_Flete.Text) = True, Txt_Flete.Text, 0)
dSeguro = IIf(IsNumeric(Txt_Seguro.Text) = True, Txt_Seguro.Text, 0)

lbl_CIF.Caption = CDbl(lbl_FOB.Caption) + dFlete + dSeguro

End Sub

Private Sub Txt_NroFP_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    Call SoloNumeros(Txt_NroFP, KeyAscii, False, 0, 8)
Else
    FunctButt1.SetFocus
End If
End Sub

Private Sub Txt_NroFP_LostFocus()
Txt_NroFP = Format(Txt_NroFP, "00000000")
End Sub

Private Sub Txt_Seguro_Change()
Dim dFlete As Double, dSeguro As Double

dFlete = IIf(IsNumeric(Txt_Flete.Text) = True, Txt_Flete.Text, 0)
dSeguro = IIf(IsNumeric(Txt_Seguro.Text) = True, Txt_Seguro.Text, 0)

lbl_CIF.Caption = CDbl(lbl_FOB.Caption) + dFlete + dSeguro
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
    End If
End Sub

Private Sub txtCod_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctButt1.SetFocus
        
    Else
        Call SoloNumeros(txtCod_OrdComp, KeyAscii, False, 0, 6)
    End If
End Sub

Private Sub txtCod_OrdComp_LostFocus()
    txtCod_OrdComp.Text = Format(Trim(txtCod_OrdComp.Text), "000000")
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
Public Sub BUSCA_CLIENTE(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    'If Trim(txtNom_Cliente.Text) <> "" Then CARGA_GRID
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(descripcion)
'                         OptCliPend.SetFocus
                         CODIGO = "": descripcion = ""
                    '     CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub
'
'Private Sub Preliminar_Docum_Ventas(tipo As Boolean)
''Copiado de TOS
'On Error GoTo errorx
'
'Dim ssql As String, Num_Corre As String, Rs As New ADODB.Recordset
'Dim aMess(4), I As Integer
'
'
' If Imprimir_FACTURA(lbl_IdproformaFactura.Caption) = False Then
'   MsgBox "Problemas de Impresion con el Documento Nr " & GridEX1.Columns("Num_Docum_Ventas"), vbInformation, "ERROR"
'   BUSCAR
'   Exit Sub
' End If
'
'Exit Sub
'Resume
'errorx:
'    ErrorHandler err, "Autoriza Documentos"
'End Sub
'
'Public Function Imprimir_FACTURA(IdProformaFactura As String) As Boolean
''Copiado de TOS
'
'Dim Rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset, strSQL As String, scnt As Integer
'scnt = 0
'With rsFactura
'
'
'        strSQL = "Ventas_Emite_Factura_Ventas '" & IdProformaFactura & "'"
'
'        Set rsFactura = CargarRecordSetDesconectado(strSQL, cConnect)
'        Call Proforma_Factura
'        scnt = 2
'
'
'        If rsFactura.RecordCount = 0 Then
'
'            Imprimir_FACTURA = False
'            Exit Function
'
'        End If
'
'
'End With
'
'
'Imprimir_FACTURA = True
'
'End Function
'
'Sub Proforma_Factura()
'On Error GoTo ErrorImpresion
'Dim oo As Object, lvSql As String, lvRuta As String
'
'    Set oo = CreateObject("excel.application")
'
'
'    oo.Workbooks.Open vRuta & "\Factura_Proforma_Exportacion.XLT"
'
'
'    oo.Visible = True
'    oo.displayalerts = False
'    oo.Run "Reporte", rsFactura
'    Set oo = Nothing
'
'    Exit Sub
'ErrorImpresion:
'    Set oo = Nothing
'    MsgBox "Hubo error en la impresion de La Factura " & err.Description, vbCritical, "Impresion"
'End Sub
'
Private Sub txtSer_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCod_OrdComp.SetFocus
    Else
        Call SoloNumeros(txtSer_OrdComp, KeyAscii, False, 0, 3)
    End If
End Sub

Private Sub txtSer_OrdComp_LostFocus()
    txtSer_OrdComp.Text = Format(Trim(txtSer_OrdComp.Text), "000")
End Sub


