VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmSolicitudDesaColoresLocal 
   Caption         =   "Solicitud de Desarrollo de Colores"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   13095
      Begin GridEX20.GridEX GridEX1 
         Height          =   5625
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   9922
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmSolicitudDesarrolloColores.frx":0000
         Column(2)       =   "FrmSolicitudDesarrolloColores.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmSolicitudDesarrolloColores.frx":016C
         FormatStyle(2)  =   "FrmSolicitudDesarrolloColores.frx":02A4
         FormatStyle(3)  =   "FrmSolicitudDesarrolloColores.frx":0354
         FormatStyle(4)  =   "FrmSolicitudDesarrolloColores.frx":0408
         FormatStyle(5)  =   "FrmSolicitudDesarrolloColores.frx":04E0
         FormatStyle(6)  =   "FrmSolicitudDesarrolloColores.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmSolicitudDesarrolloColores.frx":0678
      End
   End
   Begin VB.Frame FraBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Buscar por:"
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
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13125
      Begin VB.OptionButton OptCarta 
         BackColor       =   &H00C0FFFF&
         Caption         =   "SOLICITUD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtNum_Carta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptCliente 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   320
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Frame FraCliente 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1440
         TabIndex        =   1
         Top             =   120
         Width           =   9015
         Begin VB.CommandButton cmdBusCliente 
            Caption         =   "..."
            Height          =   285
            Left            =   1365
            TabIndex        =   4
            Tag             =   "..."
            Top             =   120
            Width           =   300
         End
         Begin VB.TextBox txtDes_Cliente 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   120
            Width           =   4485
         End
         Begin VB.TextBox txtAbr_Cliente 
            Height          =   285
            Left            =   720
            TabIndex        =   2
            Top             =   120
            Width           =   615
         End
      End
      Begin FunctionsButtons.FunctButt FBBuscar 
         Height          =   495
         Left            =   11760
         TabIndex        =   6
         Top             =   225
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1100
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "NRO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1800
         TabIndex        =   11
         Top             =   810
         Width           =   330
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   555
      Left            =   3360
      TabIndex        =   9
      Top             =   7320
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   979
      Custom          =   $"FrmSolicitudDesarrolloColores.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   530
      ControlSeparator=   70
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   5760
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmSolicitudDesaColoresLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CODIGO, Descripcion As String
Dim StrSQL As String
Dim i As Integer
Dim tmpCliente As String

Private Sub cmdBusca_Temporada_Click()
     Call BUSCA_TEMPORADA
     FBBuscar.SetFocus
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM Tx_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txtAbr_Cliente.Text = CODIGO
        txtDes_Cliente.Text = Descripcion
        'txtCod_TemCli.Enabled = True
        'txtNom_TemCli.Enabled = True
        'cmdBusca_Temporada.Enabled = True
        'txtCod_TemCli.SetFocus
        CODIGO = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub FBBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call BUSCAR
End Sub

Private Sub Form_Load()
Dim sSeguridad  As String
sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
    
'Me.FunctButt1.FunctionsUser = sSeguridad
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not oParent Is Nothing Then oParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    Load FrmAddSolicitudDesaColoresLocal
    FrmAddSolicitudDesaColoresLocal.sAccion = "I"
    FrmAddSolicitudDesaColoresLocal.TxtCod_Cliente.Text = txtAbr_Cliente.Text
    FrmAddSolicitudDesaColoresLocal.txtDes_Cliente.Text = DevuelveCampo("select nom_cliente from tg_cliente where abr_cliente ='" & txtAbr_Cliente.Text & "'", cConnect)
    'FrmAddSolicitudDesaColoresLocal.txtCod_TemCli.Text = txtCod_TemCli.Text
    'FrmAddSolicitudDesaColoresLocal.TxtDes_TemCli.Text = txtNom_TemCli.Text
    FrmAddSolicitudDesaColoresLocal.DTPSolicitud.Value = Date
    i = GridEX1.Row
    FrmAddSolicitudDesaColoresLocal.Show 1
    
    If FrmAddSolicitudDesaColoresLocal.vOk = True Then
        OptCarta.Value = True
        TxtNum_Carta.Text = FrmAddSolicitudDesaColoresLocal.Num_carta
        BUSCAR
        GridEX1.Row = GridEX1.RowCount
        VerDetalleSolicitudColoresLocal.Add = 1
        Call FunctButt1_ActionClick(0, 0, "DETALLE")
    End If
    Set FrmAddSolicitudDesaColores = Nothing
    
Case "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmAddSolicitudDesaColoresLocal
    FrmAddSolicitudDesaColoresLocal.sAccion = "U"
    FrmAddSolicitudDesaColoresLocal.TxtCorr_Carta.Text = GridEX1.Value(GridEX1.Columns("Corr_Carta").Index)
    FrmAddSolicitudDesaColoresLocal.TxtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    FrmAddSolicitudDesaColoresLocal.TxtCod_Cliente.Text = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index)
    FrmAddSolicitudDesaColoresLocal.txtDes_Cliente.Text = GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    'FrmAddSolicitudDesaColoresLocal.txtCod_TemCli.Text = GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index)
    'FrmAddSolicitudDesaColoresLocal.TxtDes_TemCli.Text = GridEX1.Value(GridEX1.Columns("Nom_TemCli").Index)
    FrmAddSolicitudDesaColoresLocal.DTPSolicitud.Value = GridEX1.Value(GridEX1.Columns("Fec_Solicitada").Index)
    FrmAddSolicitudDesaColoresLocal.TxtNum_Carta.Text = GridEX1.Value(GridEX1.Columns("Numero_Carta").Index)
    i = GridEX1.Row
    FrmAddSolicitudDesaColoresLocal.Show 1
    Set FrmAddSolicitudDesaColoresLocal = Nothing
    BUSCAR
Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmAddSolicitudDesaColoresLocal
    FrmAddSolicitudDesaColoresLocal.sAccion = "D"
    FrmAddSolicitudDesaColoresLocal.TxtCorr_Carta.Text = GridEX1.Value(GridEX1.Columns("Corr_Carta").Index)
    FrmAddSolicitudDesaColoresLocal.TxtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    FrmAddSolicitudDesaColoresLocal.TxtCod_Cliente.Text = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index)
    FrmAddSolicitudDesaColoresLocal.txtDes_Cliente.Text = GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    'FrmAddSolicitudDesaColoresLocal.txtCod_TemCli.Text = GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index)
    'FrmAddSolicitudDesaColoresLocal.TxtDes_TemCli.Text = GridEX1.Value(GridEX1.Columns("Nom_TemCli").Index)
    FrmAddSolicitudDesaColoresLocal.DTPSolicitud.Value = GridEX1.Value(GridEX1.Columns("Fec_Solicitada").Index)
    FrmAddSolicitudDesaColoresLocal.TxtNum_Carta.Text = GridEX1.Value(GridEX1.Columns("Numero_Carta").Index)
    FrmAddSolicitudDesaColoresLocal.FraDatos.Enabled = False
    i = GridEX1.Row
    FrmAddSolicitudDesaColoresLocal.Show 1
    Set FrmAddSolicitudDesaColoresLocal = Nothing
    BUSCAR
Case "IMPRIMIR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Call Reporte
Case "DETALLE"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load VerDetalleSolicitudColoresLocal
    VerDetalleSolicitudColoresLocal.TxtCorr_Carta.Text = GridEX1.Value(GridEX1.Columns("Corr_Carta").Index)
    VerDetalleSolicitudColoresLocal.TxtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    VerDetalleSolicitudColoresLocal.SCLIENTE = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    VerDetalleSolicitudColoresLocal.sTemporada = GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_TemCli").Index)
    VerDetalleSolicitudColoresLocal.CARGA_GRID
    If VerDetalleSolicitudColoresLocal.Add = 1 Then
        Call VerDetalleSolicitudColoresLocal.FunctButt1_ActionClick(0, 0, "ADICIONAR")
    End If
    i = GridEX1.Row
    VerDetalleSolicitudColoresLocal.Show 1
    Set VerDetalleSolicitudColoresLocal = Nothing
    BUSCAR
'Case "FORMATO"
'    Load FrmCambioEstado
'    FrmCambioEstado.sCorr_Carta = GridEX1.Value(GridEX1.Columns("Corr_Carta").Index)
'    FrmCambioEstado.sCliente = GridEX1.Value(GridEX1.Columns("abr_Cliente").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
'    FrmCambioEstado.sTemporada = GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_TemCli").Index)
'    FrmCambioEstado.FunctButt1.Visible = False
'    FrmCambioEstado.FunctButt2.Visible = True
'    FrmCambioEstado.CARGA_GRID
'    FrmCambioEstado.Caption = "Impresión Formato Solicitud"
'    FrmCambioEstado.Show 1
'    Set FrmCambioEstado = Nothing
    'Call Formato
'Case "ESTADO"
'    Load FrmCambioEstado
'    FrmCambioEstado.sCorr_Carta = GridEX1.Value(GridEX1.Columns("Corr_Carta").Index)
'    FrmCambioEstado.FunctButt1.Visible = True
'    FrmCambioEstado.FunctButt2.Visible = False
'    FrmCambioEstado.CARGA_GRID
'    FrmCambioEstado.Caption = "Cambio Estado Solicitud"
'    FrmCambioEstado.Show 1
'    Set FrmCambioEstado = Nothing
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub OptCarta_Click()
If OptCarta Then
    FraCliente.Visible = False
    'FraCarta.Visible = True
End If
End Sub

Private Sub optcliente_Click()
If OptCliente Then
    FraCliente.Visible = True
    'FraCarta.Visible = False
End If
End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If Trim(txtAbr_Cliente.Text) = "" Then
'            cmdBusCliente_Click
'        Else
'            txtAbr_Cliente.Text = UCase(txtAbr_Cliente.Text)
'            strSQL = "SELECT Nom_Cliente FROM Tx_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
'            txtDes_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
'            txtCod_TemCli.Enabled = True
'            txtNom_TemCli.Enabled = True
'            cmdBusca_Temporada.Enabled = True
'            txtCod_TemCli.SetFocus
'        End If


        If Trim(txtAbr_Cliente.Text) = "" Then
            Call Me.BUSCA_CLIENTE(3, txtAbr_Cliente, txtDes_Cliente)
        Else
            Call Me.BUSCA_CLIENTE(1, txtAbr_Cliente, txtDes_Cliente)
        End If
        End If

End Sub

'Private Sub txtCod_TemCli_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Trim(txtCod_TemCli.Text) = "" Then
'            Call BUSCA_TEMPORADA
'        Else
'            StrSQL = "SELECT Cod_Cliente_tex FROM Tx_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
'            StrSQL = "SELECT Nom_TemCli FROM Tx_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cConnect) & "' AND Cod_TemCli='" & txtCod_TemCli.Text & "'"
'            txtNom_TemCli.Text = DevuelveCampo(StrSQL, cConnect)
'
'            FBBuscar.SetFocus
'        End If
'    End If
'End Sub

Private Sub txtDes_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
''        If Len(txtDes_Cliente) > 4 Then
'            strSQL = "SELECT Abr_Cliente FROM Tx_CLIENTE WHERE Nom_Cliente LIKE '%" & Trim(txtDes_Cliente.Text) & "%'"
'            txtAbr_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
'            strSQL = "SELECT Nom_Cliente FROM Tx_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
'            txtDes_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
'            txtCod_TemCli.Enabled = True
'            txtNom_TemCli.Enabled = True
'            cmdBusca_Temporada.Enabled = True
'            txtCod_TemCli.SetFocus
'
'
''        Else
''            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
''            txtDes_Cliente.SetFocus
''        End If


        If Trim(txtDes_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3, txtAbr_Cliente, txtDes_Cliente)
        Else
            Call BUSCA_CLIENTE(2, txtAbr_Cliente, txtDes_Cliente)
        End If
        'txtCod_TemCli.SetFocus

    End If
End Sub


Public Sub BUSCA_CLIENTE(Tipo As Integer, CtrCodigo As TextBox, CtrDescripcion As TextBox)
    Select Case Tipo
        Case 1:
                    StrSQL = "SELECT nom_cliente FROM Tx_cliente WHERE abr_cliente = '" & Trim(CtrCodigo.Text) & "'"
                    txtDes_Cliente.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim RS As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT abr_cliente AS 'Código', nom_cliente AS 'Descripción' FROM Tx_cliente where nom_cliente like '%" & Trim(txtDes_Cliente.Text) & "%' order by abr_cliente"
                    Else
                        oTipo.SQuery = "SELECT abr_cliente AS 'Código', nom_cliente AS 'Descripción' FROM Tx_cliente order by abr_cliente"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         txtAbr_Cliente.Text = Trim(CODIGO)
                         txtDes_Cliente.Text = Trim(Descripcion)
                         CODIGO = "": Descripcion = ""

'                         Me.txtCod_TemCli.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
StrSQL = "SELECT cod_cliente_tex FROM tx_cliente WHERE abr_cliente = '" & Trim(txtAbr_Cliente.Text) & "'"
tmpCliente = DevuelveCampo(StrSQL, cConnect)

End Sub


Private Sub BUSCA_TEMPORADA()
'    Dim oTipo As New frmBusqGeneral
'    Dim RS As New ADODB.Recordset
'    Set oTipo.oParent = Me
'    StrSQL = "SELECT Cod_Cliente_tex FROM TX_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
'    oTipo.SQuery = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM Tx_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cConnect) & "'"
'    oTipo.Cargar_Datos
'    oTipo.Show 1
'    If CODIGO <> "" Then
'        txtCod_TemCli.Text = CODIGO
'        txtNom_TemCli.Text = Descripcion
'    End If
'    Set oTipo = Nothing
'    Set RS = Nothing
'
'    FBBuscar.SetFocus
End Sub

Sub BUSCAR()
Dim sopcion As Integer

If OptCliente Then
    sopcion = 0
Else
    sopcion = 1
End If

If sopcion = 0 Then
    If Trim(txtAbr_Cliente.Text) = "" Then
        MsgBox "Ingrese Cliente", vbInformation
        Exit Sub
    End If
'    If Trim(txtCod_TemCli.Text) = "" Then
'        MsgBox "Ingrese Temporada", vbInformation
'        Exit Sub
'    End If
Else
    If Trim(TxtNum_Carta.Text) = "" Then
        MsgBox "Ingrese Numero Carta a buscar", vbInformation
        Exit Sub
     End If
End If

StrSQL = "SELECT Cod_Cliente_tex FROM Tx_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"

StrSQL = "es_muestra_solicitudes_desarrollo_Local '" & DevuelveCampo(StrSQL, cConnect) & "',''," & Val(sopcion) & "," & Val(TxtNum_Carta.Text)

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)

GridEX1.Columns("Corr_Carta").Width = 0
GridEX1.Columns("Descripcion").Width = 2700
GridEX1.Columns("Fec_Creacion").Width = 1700
GridEX1.Columns("Fec_Solicitada").Width = 1150
GridEX1.Columns("Numero_Carta").Width = 900

GridEX1.Columns("cod_Temcli").Visible = False
GridEX1.Columns("nom_temcli").Visible = False


GridEX1.Columns("Numero_Carta").Caption = "Num. Carta"
GridEX1.Row = i
GridEX1.FrozenColumns = 5

End Sub

Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Cadena
    StrSQL = "SELECT Cod_Cliente_tex FROM Tx_CLIENTE WHERE Abr_Cliente='" & GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index) & "'"
    Cadena = "es_muestra_solicitudes_desarrollo_Local '" & DevuelveCampo(StrSQL, cConnect) & "','" & GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index) & "'"
    Ruta = vRuta & "\RptSolDesaColores_Local.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index), GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_TemCli").Index), Cadena, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

'Sub Formato()
'On Error GoTo hand
'Dim oo As Object
'Dim Ruta As String
'Dim Cadena As String
'Dim sCliente, sTemporada As String
'
'    sCliente = txtAbr_Cliente.Text & "-" & DevuelveCampo("select nom_cliente from tg_cliente where abr_cliente ='" & txtAbr_Cliente.Text & "'", cCONNECT)
'    sTemporada = TxtCod_TemCli.Text & "-" & txtNom_TemCli.Text
'    Cadena = "es_muestra_solicitudes_desarrollo_detalle '" & GridEX1.Value(GridEX1.Columns("Corr_Carta").Index) & "'"
'
'    Ruta = vRuta & "\RptFormato_solicitud.xlt"
'    Set oo = CreateObject("excel.application")
'    oo.Workbooks.Open Ruta
'    oo.Visible = True
'    oo.DisplayAlerts = False
'    oo.Run "Reporte", GridEX1.Value(GridEX1.Columns("Corr_Carta").Index), GridEX1.Value(GridEX1.Columns("Descripcion").Index), sCliente, sTemporada, Cadena, cCONNECT
'    Set oo = Nothing
'Exit Sub
'hand:
'    ErrorHandler Err, "GeneraReportes"
'    Set oo = Nothing
'End Sub

Private Sub TxtNum_Carta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


