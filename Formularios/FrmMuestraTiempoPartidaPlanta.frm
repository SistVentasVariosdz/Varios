VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMuestraTiempoPartidaPlanta 
   Caption         =   "TIEMPO POR AREA DE LAS PARTIDAS"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkExpandir 
      BackColor       =   &H80000010&
      Caption         =   "EXPANDIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   14040
      TabIndex        =   17
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.OptionButton Option1 
         Caption         =   "DIAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   11760
         TabIndex        =   19
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "FECHAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   11760
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Txt_Cod_Usuario 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   645
         Width           =   1095
      End
      Begin VB.TextBox Txt_DesUsuario 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   645
         Width           =   3495
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   6480
         TabIndex        =   6
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   7560
         TabIndex        =   5
         Top             =   285
         Width           =   5175
      End
      Begin VB.ComboBox CboStatus 
         Height          =   315
         Left            =   6480
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   645
         Width           =   3255
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&IMPRIMIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   75890689
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   270
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         _Version        =   393216
         Format          =   75890689
         CurrentDate     =   38182
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "VENDEDOR:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   645
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "HASTA:"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DESDE:"
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
         Left            =   480
         TabIndex        =   13
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENTE:"
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
         Left            =   5760
         TabIndex        =   12
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "STATUS:"
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
         Left            =   5760
         TabIndex        =   11
         Top             =   645
         Width           =   645
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Left            =   13560
         TabIndex        =   1
         Top             =   120
         Width           =   75
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   7365
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   12991
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      RowHeight       =   20
      AllowEdit       =   0   'False
      HeaderFontName  =   "Arial"
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   300
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmMuestraTiempoPartidaPlanta.frx":0000
      FormatStyle(2)  =   "FrmMuestraTiempoPartidaPlanta.frx":0128
      FormatStyle(3)  =   "FrmMuestraTiempoPartidaPlanta.frx":01D8
      FormatStyle(4)  =   "FrmMuestraTiempoPartidaPlanta.frx":028C
      FormatStyle(5)  =   "FrmMuestraTiempoPartidaPlanta.frx":0364
      FormatStyle(6)  =   "FrmMuestraTiempoPartidaPlanta.frx":041C
      FormatStyle(7)  =   "FrmMuestraTiempoPartidaPlanta.frx":04FC
      ImageCount      =   0
      PrinterProperties=   "FrmMuestraTiempoPartidaPlanta.frx":051C
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   360
      Top             =   6480
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMuestraTiempoPartidaPlanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String, Descripcion As String
Private StrSQL As String
Private sOpcion As String
Private indice As Integer
Private tipo As String

Private Sub chkExpandir_Click()
    If GridEX1.RowCount = 0 Then Exit Sub
    With GridEX1
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

Private Sub cmdBuscar_Click()
Call BUSCAR
End Sub
Private Sub Form_Load()
    DTPInicio = Date - 7
    DTPHasta = Date
    'cmdBuscar.SetFocus
    StrSQL = "TI_MUESTRA_STATUS_PARTIDAS"
    Call LlenaCombo(CboStatus, StrSQL, cConnect)
    If CboStatus.ListCount > 0 Then CboStatus.ListIndex = 0
    tipo = "D"
    indice = "1"

End Sub



Private Sub Option1_Click(Index As Integer)
indice = Index
Set GridEX1.ADORecordset = Nothing
If indice = 0 Then
    tipo = "F"
Else
    tipo = "D"
End If

End Sub

Private Sub Txt_Cod_Usuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_DesUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub


Public Sub Busca_Trabajador()
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
      
StrSQL = "Tg_Sm_Muestra_Operario_Caracteristica '001'"
    With frmBusqGeneralOperario
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Apellido_Paterno").Caption = "Ape Paterno"
        .DGridLista.Columns("Apellido_Paterno").Width = 1500
        .DGridLista.Columns("Apellido_Materno").Caption = "Ape Materno"
        .DGridLista.Columns("Apellido_Materno").Width = 1500
        .DGridLista.Columns("Nombre_Trabajador").Caption = "Nombres"
        .DGridLista.Columns("Nombre_Trabajador").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            Txt_Cod_Usuario = Trim(rstAux!CODIGO)
            Txt_Cod_Usuario.Tag = Left(Trim(rstAux!CODIGO), 1)
            Txt_DesUsuario = Trim(rstAux!Apellido_Paterno) + " " + Trim(rstAux!Apellido_Materno) + " " + Trim(rstAux!Nombre_Trabajador)
            Txt_DesUsuario.Tag = Right(Trim(rstAux!CODIGO), 4)
            'stip_Trabajador = Left(rstAux!codigo, 1)
            'scod_trabajador = Right(rstAux!codigo, 4)
        End If
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Color (" & Opcion & ")"
End Sub

'Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Trim(txtAbr_Cliente.Text) = "" Then
'        Call BUSCA_CLIENTE
'    Else
'        txtNom_Cliente.Text = DevuelveCampo("select nom_cliente from tx_cliente where abr_cliente='" & Trim(txtAbr_Cliente.Text) & "'", cConnect)
'    End If
'    cmdBuscar.SetFocus
'End If
'End Sub
'
'Public Sub BUSCA_CLIENTE()
'    Dim oTipo As New frmBusqGeneral
'    Dim rs As New ADODB.Recordset
'    Set oTipo.oParent = Me
'
'    oTipo.sQuery = "EXEC ti_muestra_clientes_tinto_guias"
'
'    oTipo.Cargar_Datos
'    oTipo.gexList.Columns(2).Width = 3200
'    oTipo.Show 1
'    If codigo <> "" Then
'         Me.txtAbr_Cliente.Text = Trim(codigo)
'         Me.txtNom_Cliente.Text = Trim(Descripcion)
'         Me.txtAbr_Cliente.Tag = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente ='" & Trim(codigo) & "'", cConnect)
'         codigo = "": Descripcion = ""
'    End If
'    Set oTipo = Nothing
'    Set rs = Nothing
'End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If
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



Public Sub BUSCA_CLIENTE(tipo As Integer)
    Select Case tipo
        Case 1:
                    StrSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                    If Trim(txtNom_Cliente.Text) <> "" Then BUSCAR
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.gexList.Columns(2).Width = 4850
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
                         Me.txtAbr_Cliente.Tag = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente ='" & Trim(CODIGO) & "'", cConnect)
'                         OptCliPend.SetFocus
                         CODIGO = "": Descripcion = ""
                         BUSCAR
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub



Sub BUSCAR()

  On Error GoTo fin
  
Me.txtAbr_Cliente.Tag = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente ='" & Trim(Me.txtAbr_Cliente.Text) & "'", cConnect)

StrSQL = "CN_MUESTRA_TIEMPO_TELA_PLANTA '" & DTPInicio.Value & "','" & DTPHasta.Value & "','" & Me.txtAbr_Cliente.Tag & "','" & Trim(Txt_Cod_Usuario.Tag) + Trim(Txt_DesUsuario.Tag) & "','" & Left(CboStatus, 1) & "','" & tipo & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)


'GridEX1.Columns("Cod_Almacen").Width = 700
'GridEX1.Columns("orden_compra").Width = 1050
'GridEX1.Columns("guia").Width = 1200
'GridEX1.Columns("fecha_recepcion").Width = 1100
'GridEX1.Columns("tela").Width = 2500
'GridEX1.Columns("comb").Width = 0
'GridEX1.Columns("talla").Width = 700
'GridEX1.Columns("calidad").Width = 700
'GridEX1.Columns("kgs_crudos_segun_guia").Width = 900
'GridEX1.Columns("nro_rollos_segun_guia").Width = 900
'GridEX1.Columns("kgs_pesados").Width = 800
'GridEX1.Columns("partida").Width = 700
'GridEX1.Columns("color").Width = 700
'GridEX1.Columns("nombre_color").Width = 2000
'GridEX1.Columns("nro_rollos_pesados").Width = 900
'GridEX1.Columns("kgs_asignados").Width = 800
'GridEX1.Columns("nro_rollos_asignados").Width = 900
'GridEX1.Columns("kgs_tenidos").Width = 800
'GridEX1.Columns("fecha_1er_ingreso_almacen").Width = 1000
'GridEX1.Columns("fecha_1er_despacho").Width = 1000
'GridEX1.Columns("Status").Width = 2000
'GridEX1.Columns("Ultimo_Proceso").Width = 2000
'
'GridEX1.Columns("kgs_crudos_segun_guia").Caption = "Kgs. Crudo segun Guia"
'GridEX1.Columns("nro_rollos_segun_guia").Caption = "Nro. Rollos segun Guia"
'GridEX1.Columns("kgs_pesados").Caption = "Kgs. Pesados"
'GridEX1.Columns("orden_compra").Caption = "Ord. Compra"
'GridEX1.Columns("Cod_Almacen").Caption = "Almacen"
'GridEX1.Columns("fecha_recepcion").Caption = "Fec. Recep."
'GridEX1.Columns("nombre_color").Caption = "Nombre Color"
'GridEX1.Columns("nro_rollos_pesados").Caption = "Nro. Rollos Pesados"
'GridEX1.Columns("kgs_asignados").Caption = "Kgs. Asignados"
'GridEX1.Columns("nro_rollos_asignados").Caption = "Nro. Rollos Asignados"
'GridEX1.Columns("kgs_tenidos").Caption = "Kgs. Tenidos"
'GridEX1.Columns("fecha_1er_ingreso_almacen").Caption = "Fecha 1er Ing. Almacen"
'GridEX1.Columns("fecha_1er_despacho").Caption = "Fecha 1er Despacho"

  Call CONFIGURA_GRILLA
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption

End Sub
Private Sub CONFIGURA_GRILLA()

    On Error GoTo SALTO_ERROR
    
    Dim C As Integer
        
    'GridEX1.FrozenColumns = 6
    With GridEX1
'        .Columns("CENTRO_COSTO").Visible = False
'        .Columns("CODIGO").Visible = False
'        .Columns("TRABAJADOR").Visible = False
'        .Columns("DES_TRABAJADOR").Visible = False
'        .Columns("DNI").Visible = False
'        .Columns("INGRESO").Visible = False
'        .Columns("CESE").Visible = False
'        .Columns("HORARIO").Visible = False
'        .Columns("TIPO").Visible = False
'        .Columns("SECUENCIA").Visible = False
'        .Columns("FEC_REGISTRO").Visible = False
'        .Columns("REGISTRO").Visible = False
'        .Columns("ES_FERIADO").Visible = False

'        For C = 1 To .Columns.Count
'            .Columns(C).HeaderAlignment = jgexAlignCenter
'            .Columns(C).TextAlignment = jgexAlignCenter
'        Next C
'
'        With .Columns("REGISTRO_CAD")
'            .Caption = "REGISTRO"
'            .Width = 1000
'        End With

        
        Dim oGroup01 As GridEX20.JSGroup
        Dim oGroup02 As GridEX20.JSGroup
        
        Set oGroup01 = .Groups.Add(.Columns("CLIENTE").Index, jgexSortAscending)
        'Set oGroup02 = .Groups.Add(.Columns("DES_TRABAJADOR").Index, jgexSortAscending)
        
        .BackColorRowGroup = &H8000000F
        If CBool(chkExpandir.Value) = True Then
            .DefaultGroupMode = jgexDGMExpanded
        Else
            .DefaultGroupMode = jgexDGMCollapsed
        End If
        .ForeColorRowGroup = vbBlue
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
        
'        Dim colHORAS As JSColumn
'
'        .GroupFooterStyle = jgexTotalsGroupFooter
'        Set colHORAS = .Columns("HORAS")
'        With colHORAS
'            .AggregateFunction = jgexSum
'            .TotalRowPrefix = ""
'        End With
        
        .SetFocus
    End With
    
    Exit Sub
    
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
    
End Sub

Private Sub CmdImprimir_Click()
If GridEX1.RowCount <= 0 Then Exit Sub

 Select Case tipo
 Case "F"
    Call Reporte
 Case "D"
    Call Reportedias
 End Select
 
End Sub
Private Sub Reporte()
Dim oo As Object
Dim sRutaLogo  As String
Dim Ruta As String
On Error GoTo errReporte

Ruta = vRuta & "\RptMuestraTiempoPartidaPlanta.xlt"

Set oo = CreateObject("excel.application")

StrSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
sRutaLogo = DevuelveCampo(StrSQL, cConnect)
    
oo.Workbooks.Open Ruta
oo.Visible = False
oo.DisplayAlerts = False
oo.Run "Reporte", sRutaLogo, GridEX1.ADORecordset, txtNom_Cliente.Text, Txt_DesUsuario.Text, CStr(DTPInicio.Value) + " - " + CStr(DTPHasta.Value)
oo.Visible = True

Set oo = Nothing

Exit Sub
errReporte:
    MsgBox "Hubo error en la impresion del Reporte de Tiempos " & err.Description, vbCritical, "Impresion"
End Sub
Private Sub Reportedias()
Dim oo As Object
Dim sRutaLogo  As String
Dim Ruta As String
On Error GoTo errReporte

Ruta = vRuta & "\RptMuestraTiempoPartidaPlantaDias.xlt"

Set oo = CreateObject("excel.application")

StrSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
sRutaLogo = DevuelveCampo(StrSQL, cConnect)
    
oo.Workbooks.Open Ruta
oo.Visible = False
oo.DisplayAlerts = False
oo.Run "Reporte", sRutaLogo, GridEX1.ADORecordset, txtNom_Cliente.Text, Txt_DesUsuario.Text, CStr(DTPInicio.Value) + " - " + CStr(DTPHasta.Value)
oo.Visible = True

Set oo = Nothing

Exit Sub
errReporte:
    MsgBox "Hubo error en la impresion del Reporte de Tiempos " & err.Description, vbCritical, "Impresion"
End Sub

'Sub Reporte()
'On Error GoTo ErrorImpresion
'Dim oo As Object
'
'    Set oo = CreateObject("excel.application")
'    oo.workbooks.Open vRuta & "\RptStatusGuiaPartidasClientes.XLT"
'    oo.Visible = True
'    'oo.run "reporte", sOpcion, Me.txtAbr_Cliente.Tag, Trim(txtNom_Cliente.Text), DTPInicio.Value, DTPHasta.Value, cConnect, "", "", Trim(Txt_Cod_Usuario.Text)
'    Set oo = Nothing
'
'    Exit Sub
'ErrorImpresion:
'    Set oo = Nothing
'    MsgBox "Hubo error en la impresion del Reporte de Guia de Remisión " & Err.Description, vbCritical, "Impresion"
'End Sub

Private Sub GridEX1_GroupByBoxHeaderClick(ByVal Group As JSGroup)
    Group.SortOrder = -Group.SortOrder
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    If GridEX1.RowCount = 0 Then Exit Sub
    Dim fmtConTipoRegistro As JSFmtCondition

    Set fmtConTipoRegistro = GridEX1.FmtConditions.Add(GridEX1.Columns("KGS_CRUDO").Index, jgexEqual, "0")

    With fmtConTipoRegistro.FormatStyle
        .ForeColor = &H8000&
        .FontSize = 8
        .BackColor = &H80000018 'vbYellow
    End With
End Sub


