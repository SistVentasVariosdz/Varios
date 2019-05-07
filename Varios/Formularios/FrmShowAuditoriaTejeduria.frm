VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmShowAuditoriaTejeduria 
   Caption         =   "Auditoria Tejeduria"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraRangoFechas 
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9480
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker DTPFecha1 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   72351745
         CurrentDate     =   38384
      End
      Begin MSComCtl2.DTPicker DTPFecha2 
         Height          =   285
         Left            =   1800
         TabIndex        =   38
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   72351745
         CurrentDate     =   38384
      End
      Begin FunctionsButtons.FunctButt FunctButt5 
         Height          =   510
         Left            =   720
         TabIndex        =   59
         Top             =   1200
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmShowAuditoriaTejeduria.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   840
         TabIndex        =   40
         Top             =   820
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   840
         TabIndex        =   39
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame FraImpresion 
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton OptListadodet 
            Caption         =   "Listado detallado x rollo auditado"
            Height          =   195
            Left            =   480
            TabIndex        =   55
            Top             =   1140
            Width           =   2655
         End
         Begin VB.OptionButton OptPrimeras 
            Caption         =   "Defectos en 1ras / 2das"
            Height          =   195
            Left            =   480
            TabIndex        =   45
            Top             =   890
            Width           =   2775
         End
         Begin VB.OptionButton OptDefectos 
            Caption         =   "Defectos que originan Calidad 3"
            Height          =   195
            Left            =   480
            TabIndex        =   41
            Top             =   650
            Width           =   2775
         End
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado por Rollo Auditado"
            Height          =   195
            Left            =   480
            TabIndex        =   27
            Top             =   390
            Width           =   2535
         End
         Begin VB.OptionButton OptResumido 
            Caption         =   "Resumido por Dia"
            Height          =   195
            Left            =   480
            TabIndex        =   26
            Top             =   150
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame FraResumido 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   3495
         Begin MSComCtl2.DTPicker DTPInicio 
            Height          =   255
            Left            =   1440
            TabIndex        =   29
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38416
         End
         Begin MSComCtl2.DTPicker DTPFin 
            Height          =   255
            Left            =   1440
            TabIndex        =   30
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38416
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   360
            TabIndex        =   32
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   360
            TabIndex        =   31
            Top             =   555
            Width           =   705
         End
      End
      Begin VB.Frame FraDetallado 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   3495
         Begin MSComCtl2.DTPicker DTFecDetallado 
            Height          =   255
            Left            =   1560
            TabIndex        =   34
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38416
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Auditoria"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   1110
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   600
         TabIndex        =   57
         Top             =   2520
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmShowAuditoriaTejeduria.frx":0099
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.Frame FraTela 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1920
         TabIndex        =   49
         Top             =   1490
         Width           =   8775
         Begin VB.TextBox TxtDes_Tela 
            Height          =   285
            Left            =   1440
            TabIndex        =   51
            Top             =   60
            Width           =   3975
         End
         Begin VB.TextBox Txtcod_Tela 
            Height          =   285
            Left            =   240
            TabIndex        =   50
            Top             =   60
            Width           =   1170
         End
         Begin MSComCtl2.DTPicker DTPInicio_Tela 
            Height          =   285
            Left            =   5640
            TabIndex        =   52
            Top             =   50
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38384
         End
         Begin MSComCtl2.DTPicker DTPFin_Tela 
            Height          =   285
            Left            =   7200
            TabIndex        =   53
            Top             =   50
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38384
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Left            =   6960
            TabIndex        =   54
            Top             =   90
            Width           =   120
         End
      End
      Begin VB.OptionButton OptTela 
         Caption         =   "Tela"
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TxtCod_OrdTra 
         Height          =   285
         Left            =   2160
         TabIndex        =   47
         Top             =   1200
         Width           =   1140
      End
      Begin VB.OptionButton OptxOT 
         Caption         =   "OT"
         Height          =   375
         Left            =   360
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame FraOtMaquina 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   8535
         Begin VB.TextBox TxtCod_MaquinaBus 
            Height          =   285
            Left            =   240
            TabIndex        =   24
            Top             =   0
            Width           =   780
         End
         Begin VB.TextBox TxtDes_MaquinaBus 
            Height          =   285
            Left            =   1080
            TabIndex        =   23
            Top             =   0
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker DTPMaquinaDesde 
            Height          =   285
            Left            =   5280
            TabIndex        =   42
            Top             =   0
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38384
         End
         Begin MSComCtl2.DTPicker DTPMaquinaHasta 
            Height          =   285
            Left            =   6840
            TabIndex        =   43
            Top             =   0
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38384
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Left            =   6600
            TabIndex        =   44
            Top             =   90
            Width           =   120
         End
      End
      Begin VB.OptionButton OptOtMaquina 
         Caption         =   "Maquina"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.Frame FraRango 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   3495
         Begin MSComCtl2.DTPicker DTPInicio_Rango 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   50
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38384
         End
         Begin MSComCtl2.DTPicker DTPFin_Rango 
            Height          =   285
            Left            =   1920
            TabIndex        =   19
            Top             =   50
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            Format          =   72351745
            CurrentDate     =   38384
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Left            =   1560
            TabIndex        =   20
            Top             =   140
            Width           =   120
         End
      End
      Begin VB.OptionButton OptRango 
         Caption         =   "Rango Fechas"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   495
         Width           =   1455
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Fecha Auditoria"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   165
         Value           =   -1  'True
         Width           =   1455
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   540
         Left            =   11760
         TabIndex        =   1
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   953
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1200
         ControlHeigth   =   520
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   72351745
         CurrentDate     =   38384
      End
   End
   Begin VB.Frame FraInspeccion 
      Caption         =   "Inspección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4680
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox TxtNom_Fabrica 
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "HIALPESA"
         Top             =   1080
         Width           =   3420
      End
      Begin VB.TextBox TxtCod_OrdPro 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox TxtCod_Fabrica 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Text            =   "001"
         Top             =   1080
         Width           =   660
      End
      Begin VB.TextBox txtCod_Maquina 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   660
      End
      Begin VB.TextBox txtDes_Maquina 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxtOT 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptNP 
         Caption         =   "NP"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton OptMaquina 
         Caption         =   "Maquina"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton OptOT 
         Caption         =   "OT"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin FunctionsButtons.FunctButt FunctButt4 
         Height          =   510
         Left            =   1080
         TabIndex        =   58
         Top             =   1800
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmShowAuditoriaTejeduria.frx":0132
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5295
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   9340
      Version         =   "2.0"
      PreviewRowIndent=   0
      AutomaticSort   =   -1  'True
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
      FormatStyle(1)  =   "FrmShowAuditoriaTejeduria.frx":01CB
      FormatStyle(2)  =   "FrmShowAuditoriaTejeduria.frx":0303
      FormatStyle(3)  =   "FrmShowAuditoriaTejeduria.frx":03B3
      FormatStyle(4)  =   "FrmShowAuditoriaTejeduria.frx":0467
      FormatStyle(5)  =   "FrmShowAuditoriaTejeduria.frx":053F
      FormatStyle(6)  =   "FrmShowAuditoriaTejeduria.frx":05F7
      FormatStyle(7)  =   "FrmShowAuditoriaTejeduria.frx":06D7
      ImageCount      =   0
      PrinterProperties=   "FrmShowAuditoriaTejeduria.frx":06F7
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   525
      Left            =   0
      TabIndex        =   56
      Top             =   7440
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   926
      Custom          =   $"FrmShowAuditoriaTejeduria.frx":08CF
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   500
      ControlSeparator=   50
   End
   Begin FunctionsButtons.FunctButt FunctButt6 
      Height          =   510
      Left            =   7920
      TabIndex        =   60
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   5160
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmShowAuditoriaTejeduria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim strSQL As String
Public CODIGO As String, Descripcion As String, TipoAdd As String
Public fila_seleccionada As Long
Dim tipo As String

Private Sub Form_Load()
DTPFecha.Value = Date
DTPInicio_Rango.Value = Date
DTPFin_Rango.Value = Date
DTPInicio_Tela.Value = Date
DTPFin_Tela.Value = Date
DTPMaquinaDesde.Value = Date
DTPMaquinaHasta.Value = Date

Call CARGA_GRID

FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'If Not oParent Is Nothing Then oParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim iHora As Integer

Select Case ActionName
Case "ADICIONAR"
    Load FrmMante_CC_Tejeduria
    With FrmMante_CC_Tejeduria
        Set .oParent = Me
        .sAccion = "I"
        .DTPFecha.Value = DTPFecha.Value
        .TxtTip_Auditor = GridEX1.Value(GridEX1.Columns("Tip_Auditor_cc").Index)
        .TxtCod_Auditor = GridEX1.Value(GridEX1.Columns("cod_Auditor_cc").Index)
        .TxtNom_Auditor = GridEX1.Value(GridEX1.Columns("Nom_Auditor").Index)
        .TxtTurno = GridEX1.Value(GridEX1.Columns("Turno").Index)
        '.dtpFecha = GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index)
        .TxtRestriccion = "C"
        .BUSCARESTRICCION ("1")
'        If Not (IsNull(GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index)) Or Trim(GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index)) = "") Then
'            .dtpFecha = GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index)
'        End If
        .DTPFecha = DevuelveCampo("select dbo.Devuelve_Dia_Inicio_Auditoria_Tejeduria ('" & Format(Now, "dd/mm/yyyy hh:mm") & "')", cConnect)
        
        '-----NUEVO---------------
        If vemp = "07" Then
            iHora = DevuelveCampo("SELECT DATEPART(HH,GETDATE())", cConnect)
            
            If iHora >= 8 And iHora <= 16 Then
                .TxtTurno.Text = "1"
            Else
                .TxtTurno.Text = "2"
            End If
            
            .DTPFecha.Enabled = False
            
        End If
        '---------------------------
        
        .Show 1
        If .vOk = True Then
            GridEX1.Row = GridEX1.RowCount
            Call FunctButt2_ActionClick(0, 0, "DETALLE")
        End If
        

        
    End With
    Set FrmMante_CC_Tejeduria = Nothing
Case "MODIFICAR"
    If GridEX1.RowCount <= 0 Then Exit Sub
    Load FrmMante_CC_Tejeduria
    With FrmMante_CC_Tejeduria
        Set .oParent = Me
        .sAccion = "U"
        .txtCod_Maquina.Enabled = False
        .txtDes_Maquina.Enabled = False
        .TxtCodigo_Rollo.Enabled = False
        .txtCod_Maquina = GridEX1.Value(GridEX1.Columns("Prefijo_Maquina").Index)
        .txtDes_Maquina = GridEX1.Value(GridEX1.Columns("Maquina").Index)
        .TxtCodigo_Rollo = GridEX1.Value(GridEX1.Columns("Codigo_Rollo").Index)
        .LblOT = GridEX1.Value(GridEX1.Columns("OT").Index)
        .LblKilos = GridEX1.Value(GridEX1.Columns("kilos").Index)
        .TxtCod_Calidad = GridEX1.Value(GridEX1.Columns("cod_Calidad").Index)
        .TxtDes_Calidad = GridEX1.Value(GridEX1.Columns("Des_Calidad").Index)
        .TxtTip_Auditor = GridEX1.Value(GridEX1.Columns("Tip_Auditor_cc").Index)
        .TxtCod_Auditor = GridEX1.Value(GridEX1.Columns("cod_Auditor_cc").Index)
        .TxtNom_Auditor = GridEX1.Value(GridEX1.Columns("Nom_Auditor").Index)
        .TxtMerma = GridEX1.Value(GridEX1.Columns("Merma").Index)
        .TxtObservaciones = GridEX1.Value(GridEX1.Columns("Observaciones").Index)
        .TxtTurno = GridEX1.Value(GridEX1.Columns("Turno").Index)
         
         If GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index) = Null Then
           .DTPFecha = Now()
         Else
           .DTPFecha = GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index)
         End If
         
        
        .TxtRestriccion.Text = Trim(GridEX1.Value(GridEX1.Columns("Cod_Restriccion").Index))
        .TxtDes_Restriccion.Text = Trim(GridEX1.Value(GridEX1.Columns("Des_Restriccion").Index))
        .TxtTip_Tejedor.Text = Trim(GridEX1.Value(GridEX1.Columns("Tip_trabajador_tejedor").Index))
        .TxtCod_Tejedor.Text = Trim(GridEX1.Value(GridEX1.Columns("cod_trabajador_tejedor").Index))
        .TxtNom_Tejedor.Text = Trim(GridEX1.Value(GridEX1.Columns("Nombre_tejedor").Index))
        .Show 1
    End With
    Set FrmMante_CC_Tejeduria = Nothing
Case "ELIMINAR"
If GridEX1.RowCount <= 0 Then Exit Sub
    Load FrmMante_CC_Tejeduria
    With FrmMante_CC_Tejeduria
        
        Set .oParent = Me
        .FraDatos.Enabled = False
        .sAccion = "D"
        .txtCod_Maquina.Enabled = False
        .txtDes_Maquina.Enabled = False
        .TxtCodigo_Rollo.Enabled = False
        .txtCod_Maquina = GridEX1.Value(GridEX1.Columns("Prefijo_Maquina").Index)
        .txtDes_Maquina = GridEX1.Value(GridEX1.Columns("Maquina").Index)
        .TxtCodigo_Rollo = GridEX1.Value(GridEX1.Columns("Codigo_Rollo").Index)
        .LblOT = GridEX1.Value(GridEX1.Columns("OT").Index)
        .LblKilos = GridEX1.Value(GridEX1.Columns("kilos").Index)
        .TxtCod_Calidad = GridEX1.Value(GridEX1.Columns("cod_Calidad").Index)
        .TxtDes_Calidad = GridEX1.Value(GridEX1.Columns("Des_Calidad").Index)
        .TxtTip_Auditor = GridEX1.Value(GridEX1.Columns("Tip_Auditor_cc").Index)
        .TxtCod_Auditor = GridEX1.Value(GridEX1.Columns("cod_Auditor_cc").Index)
        .TxtNom_Auditor = GridEX1.Value(GridEX1.Columns("Nom_Auditor").Index)
        .TxtMerma = GridEX1.Value(GridEX1.Columns("Merma").Index)
        .TxtObservaciones = GridEX1.Value(GridEX1.Columns("Observaciones").Index)
        .TxtTurno = GridEX1.Value(GridEX1.Columns("Turno").Index)
        .DTPFecha = GridEX1.Value(GridEX1.Columns("Fec_auditoria").Index)
        .TxtRestriccion.Text = Trim(GridEX1.Value(GridEX1.Columns("Cod_Restriccion").Index))
        .TxtDes_Restriccion.Text = Trim(GridEX1.Value(GridEX1.Columns("Des_Restriccion").Index))
        .TxtTip_Tejedor.Text = Trim(GridEX1.Value(GridEX1.Columns("Tip_trabajador_tejedor").Index))
        .TxtCod_Tejedor.Text = Trim(GridEX1.Value(GridEX1.Columns("cod_trabajador_tejedor").Index))
        .TxtNom_Tejedor.Text = Trim(GridEX1.Value(GridEX1.Columns("Nombre_tejedor").Index))
        .Show 1
    End With
    Set FrmMante_CC_Tejeduria = Nothing
Case "DETALLE"
    If GridEX1.RowCount <= 0 Then Exit Sub
    Load FrmShowAuditoriaTejeduria_Detalle
    FrmShowAuditoriaTejeduria_Detalle.Caption = FrmShowAuditoriaTejeduria_Detalle.Caption & Space(1) & "Maq. :" & GridEX1.Value(GridEX1.Columns("Prefijo_Maquina").Index) & " - " & " Rollo : " & GridEX1.Value(GridEX1.Columns("Codigo_Rollo").Index)
    FrmShowAuditoriaTejeduria_Detalle.sPrefijo_Maquina = GridEX1.Value(GridEX1.Columns("Prefijo_Maquina").Index)
    FrmShowAuditoriaTejeduria_Detalle.sCodigo_Rollo = GridEX1.Value(GridEX1.Columns("Codigo_Rollo").Index)
    FrmShowAuditoriaTejeduria_Detalle.sOT = GridEX1.Value(GridEX1.Columns("OT").Index)
    FrmShowAuditoriaTejeduria_Detalle.BUSCAR
    FrmShowAuditoriaTejeduria_Detalle.Show 1
    Set FrmShowAuditoriaTejeduria_Detalle = Nothing
Case "IMPRIMIR"
    DTPInicio.Value = "01/" & Month(DTPFecha.Value) & "/" & Year(DTPFecha.Value)
    DTPFin.Value = DevuelveCampo("select dbo.tg_obtiene_dia_ultimo_ano_mes(" & Year(DTPFecha.Value) & "," & Month(DTPFecha.Value) & ")", cConnect)
    DTFecDetallado.Value = Date - 1
    FraImpresion.Visible = True
Case "INSPECCION"
    Load FrmReporteAuditoriaTejeduriaRollos
    FrmReporteAuditoriaTejeduriaRollos.DTPInicio.Value = Me.DTPInicio_Rango.Value
    FrmReporteAuditoriaTejeduriaRollos.DTPFin.Value = DTPFin_Rango.Value
    FrmReporteAuditoriaTejeduriaRollos.Show vbModal
    Set FrmReporteAuditoriaTejeduriaRollos = Nothing
Case "ROLLOS"
    tipo = "1"
    DTPFecha1.Value = DTPInicio_Rango.Value
    DTPFecha2.Value = DTPFin_Rango.Value
    strSQL = "cc_rollos_tejeduria_por_Fecha_rango "
    FraRangoFechas.Visible = True
Case "INSPECTOR"
    tipo = "2"
    strSQL = "cc_rollos_tejeduria_por_auditor_rango "
    DTPFecha1.Value = DTPInicio_Rango.Value
    DTPFecha2.Value = DTPFin_Rango.Value
    FraRangoFechas.Visible = True
Case "VISTA"
    Call Reporte_Vista
Case "POROT"
    Call Reporte_porOT
Case "TELA"
    Call Reporte_porTela
    
Case "CALIDAD3-4"
    Load frmReporteRollosCalidad3_4
    frmReporteRollosCalidad3_4.DTPInicio.Value = Me.DTPInicio_Rango.Value
    frmReporteRollosCalidad3_4.DTPFin.Value = DTPFin_Rango.Value
    frmReporteRollosCalidad3_4.Show vbModal
    Set frmReporteRollosCalidad3_4 = Nothing
    
End Select
End Sub


Sub CARGA_GRID()
Dim sOpcion As String

If OptFecha Then
    sOpcion = "1"
ElseIf OptRango Then
    sOpcion = "2"
    If DevuelveCampo("select datediff(dd,'" & DTPInicio_Rango & "','" & DTPFin_Rango & "')", cConnect) > 45 Then
        MsgBox "Rango no debe exceder los 45 dias, verifique", vbCritical
        Exit Sub
    End If
ElseIf OptOtMaquina Then
    sOpcion = "3"
    If DevuelveCampo("select datediff(dd,'" & DTPMaquinaDesde & "','" & DTPMaquinaHasta & "')", cConnect) > 45 Then
        MsgBox "Rango no debe exceder los 45 dias, Verifique", vbCritical
        Exit Sub
    End If
ElseIf OptxOT Then
    sOpcion = "4"
ElseIf OptTela Then
    sOpcion = "5"
    If DevuelveCampo("select datediff(mm,'" & DTPInicio_Tela & "','" & DTPFin_Tela & "')", cConnect) > 6 Then
        MsgBox "Rango no debe exceder los 6 meses, Verifique", vbCritical
        Exit Sub
    End If
End If

strSQL = "CC_MUESTRA_AUDITORIA_TEJEDURIA_ROLLOS_NUEVO '" & sOpcion & "','" & Format(DTPFecha.Value, "dd/mm/yyyy") & "','" & Format(DTPInicio_Rango.Value, "dd/mm/yyyy") & "','" & Format(DTPFin_Rango.Value, "dd/mm/yyyy") & "','" & Trim(TxtCod_MaquinaBus.Text) & "','" & DTPMaquinaDesde.Value & "','" & DTPMaquinaHasta.Value & "','" & TxtCod_OrdTra.Text & "','" & Txtcod_Tela.Text & "','" & Format(DTPInicio_Tela.Value, "dd/mm/yyyy") & "','" & Format(DTPFin_Tela.Value, "dd/mm/yyyy") & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

GridEX1.Columns("Fec_auditoria").Width = 1100
GridEX1.Columns("Prefijo_maquina").Width = 650
GridEX1.Columns("Prefijo_maquina").Caption = "Maq."
GridEX1.Columns("Maquina").Width = 2000
GridEX1.Columns("Codigo_Rollo").Width = 1000
GridEX1.Columns("OT").Width = 800
GridEX1.Columns("Tip_auditor_cc").Width = 0
GridEX1.Columns("Cod_Auditor_cc").Width = 0
GridEX1.Columns("Nom_Auditor").Width = 2700
GridEX1.Columns("Kilos").Width = 850
GridEX1.Columns("Merma").Width = 850
GridEX1.Columns("Turno").Width = 700
GridEX1.Columns("Cod_Restriccion").Width = 800
GridEX1.Columns("Des_Restriccion").Width = 1800
GridEX1.Columns("Tip_Trabajador_tejedor").Width = 0
GridEX1.Columns("Cod_trabajador_tejedor").Width = 0
GridEX1.Columns("Nombre_tejedor").Width = 2000
GridEX1.Columns("Creacion").Width = 1800

GridEX1.Columns("Cod_Calidad").Width = 750

GridEX1.FrozenColumns = 5

GridEX1.MoveLast
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Mensaje As Variant
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "CANCELAR"
    FraImpresion.Visible = False
End Select
End Sub

Sub Reporte()
Dim Mes_Ano As String, strSQL1 As String, cadena As String

On Error GoTo ErrorImpresion
Dim oo As Object

    cadena = "Del " & DTPInicio.Value & " al " & DTPFin.Value

    Set oo = CreateObject("excel.application")

    If OptResumido Then
        strSQL = "cc_reporte_tejeduria_resumen_diario_NEW '" & DTPInicio.Value & "','" & DTPFin.Value & "'"
        oo.Workbooks.Open vRuta & "\RptAuditoriaTejeduria.XLT"
    ElseIf OptDetallado Then
        strSQL = "cc_reporte_tejeduria_detallado_x_rollo_new '" & DTFecDetallado.Value & "','" & DTFecDetallado.Value & "'"
        oo.Workbooks.Open vRuta & "\RptAuditoriaTejeduriaDetalladoxRollo.XLT"
    ElseIf OptDefectos Then
        strSQL = "cc_muestra_Motivos_Rollos_3raCalidad '" & DTPInicio.Value & "','" & DTPFin.Value & "'"
        strSQL1 = "cc_muestra_motivos_todo_rollo '" & DTPInicio.Value & "','" & DTPFin.Value & "'"
    
        oo.Workbooks.Open vRuta & "\RptDetalleDefectosxRolloCalidad3.XLT"
    ElseIf OptListadodet Then
    
        strSQL = "cc_reporte_tejeduria_detallado_x_rollo_new2 '" & DTPInicio.Value & "','" & DTPFin.Value & "'"
        oo.Workbooks.Open vRuta & "\RptListadodetalladoxRolloauditado.XLT"
    
 
    Else
        strSQL = "cc_muestra_Motivos_Rollos_Primeras  '" & DTPInicio.Value & "','" & DTPFin.Value & "'"
        strSQL1 = "cc_muestra_motivos_todo_rollo_1ras  '" & DTPInicio.Value & "','" & DTPFin.Value & "'"
        
        oo.Workbooks.Open vRuta & "\RptDetalleDefectosxRollo1ras.XLT"
    End If
    
    
    oo.Visible = True
    oo.DisplayAlerts = False
    
    If OptResumido Then
        oo.Run "reporte", DTPInicio.Value, DTPFin.Value, strSQL, cConnect
    ElseIf OptDetallado Then
        oo.Run "reporte", DTFecDetallado.Value, DTFecDetallado.Value, strSQL, cConnect
    ElseIf OptListadodet Then
         oo.Run "reporte", strSQL, cConnect, DTPInicio.Value, DTPFin.Value
    Else
        oo.Run "reporte", strSQL, strSQL1, cadena, cConnect
    End If
    Set oo = Nothing
    
    FraImpresion.Visible = False
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Private Sub FunctButt5_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte_Rango
Case "CANCELAR"
    FraRangoFechas.Visible = False
End Select
End Sub

Private Sub FunctButt6_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub OptDefectos_Click()
FraResumido.Visible = True
FraDetallado.Visible = False
End Sub

Private Sub OptDetallado_Click()
FraResumido.Visible = False
FraDetallado.Visible = True
End Sub

Private Sub OptFecha_Click()
DTPFecha.Visible = True
FraRango.Visible = False
FraOtMaquina.Visible = False
TxtCod_OrdTra.Visible = False
End Sub

Private Sub Option2_Click()
FraResumido.Visible = False
FraDetallado.Visible = True
End Sub


Private Sub OptListadodet_Click()
FraResumido.Visible = True
FraDetallado.Visible = False
End Sub

Private Sub OptOtMaquina_Click()
FraOtMaquina.Visible = True
DTPFecha.Visible = False
FraRango.Visible = False
TxtCod_MaquinaBus.SetFocus
TxtCod_OrdTra.Visible = False
End Sub

Private Sub OptRango_Click()
FraRango.Visible = True
DTPFecha.Visible = False
FraOtMaquina.Visible = False
TxtCod_OrdTra.Visible = False
End Sub


Private Sub BUSCAMAQUINA(Opcion As Integer)
On Error GoTo Fin
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Prefijo_Maquina, Des_Maquina_Tejeduria FROM TX_MAQUINAS_TEJEDURIA WHERE "
    txtCod_Maquina = Trim(txtCod_Maquina)
    txtDes_Maquina = Trim(txtDes_Maquina)
    Select Case Opcion
    Case 1: strSQL = strSQL & "Prefijo_Maquina like '%" & TxtCod_MaquinaBus & "%'"
    Case 2: strSQL = strSQL & "Des_Maquina_Tejeduria like '%" & TxtDes_MaquinaBus & "%'"
    End Select
    TxtCod_MaquinaBus = ""
    TxtDes_MaquinaBus = ""
    fila_seleccionada = 0
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        .gexList.Columns("Prefijo_Maquina").Caption = "Prefijo"
        .gexList.Columns("Prefijo_Maquina").Width = 1000
        .gexList.Columns("Des_Maquina_Tejeduria").Caption = "Máquina"
        .gexList.Columns("Des_Maquina_Tejeduria").Width = 5000
        
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
            rstAux.AbsolutePosition = fila_seleccionada
            TxtCod_MaquinaBus = Trim(rstAux!prefijo_maquina)
            TxtDes_MaquinaBus = Trim(rstAux!Des_Maquina_Tejeduria)
            FunctButt1.SetFocus
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Maquina (" & Opcion & ")"
End Sub

Private Sub OptResumido_Click()
FraResumido.Visible = True
FraDetallado.Visible = False
End Sub

Private Sub OptTela_Click()
FraOtMaquina.Visible = True
DTPFecha.Visible = False
FraRango.Visible = False
TxtCod_OrdTra.Visible = False
FraTela.Visible = True
Txtcod_Tela.SetFocus
End Sub

Private Sub OptxOT_Click()
FraOtMaquina.Visible = True
DTPFecha.Visible = False
FraRango.Visible = False
TxtCod_OrdTra.Visible = True
TxtCod_OrdTra.SetFocus
End Sub

Private Sub TxtCod_MaquinaBus_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCAMAQUINA 1
End Sub

Private Sub TxtCod_OrdTra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCod_OrdTra.Text = Format(TxtCod_OrdTra, "0000")
    strSQL = "select count(*) from tx_ordtra where cod_tipordtra = 'Tj' and cod_ordtra='" & Trim(TxtCod_OrdTra.Text) & "'"
    If DevuelveCampo(strSQL, cConnect) = 0 Then
        MsgBox "La Orden de Trabajo no Existe", vbCritical, "Orden de Trabajo"
        TxtCod_OrdTra.SetFocus
        SelectionText TxtCod_OrdTra
    Else
        FunctButt1.SetFocus
    End If
End If
End Sub

Private Sub txtcod_tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Txtcod_Tela.Text) = "" Then
        MsgBox ("Sirvase ingresar un codigo de Item")
    Else
        Txtcod_Tela.Text = CompletaCodigo(Trim(Txtcod_Tela.Text), 8, 2)
        strSQL = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela='" & Txtcod_Tela.Text & "'"
        TxtDes_Tela.Text = DevuelveCampo(strSQL, cConnect)
        FunctButt1.SetFocus
    End If
End If
End Sub

Private Sub TxtDes_MaquinaBus_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCAMAQUINA 2
End Sub

Sub Reporte_Rango()
Dim Mes_Ano As String

On Error GoTo ErrorImpresion
Dim oo As Object

    Set oo = CreateObject("excel.application")

    If tipo = 1 Then
        strSQL = "cc_rollos_tejeduria_por_Fecha_rango '" & DTPFecha1.Value & "','" & DTPFecha2.Value & "'"
        oo.Workbooks.Open vRuta & "\RptRollosProd_e_InspecporRango.xlt"
    ElseIf tipo = 2 Then
        strSQL = "cc_rollos_tejeduria_por_auditor_rango '" & DTPFecha1.Value & "','" & DTPFecha2.Value & "'"
        oo.Workbooks.Open vRuta & "\RptRollos_x_InspectorporRango.XLT"
    End If
    
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", "Del " & DTPFecha1.Value & " Al " & DTPFecha2.Value, strSQL, cConnect
    Set oo = Nothing
    
    FraRangoFechas.Visible = False
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Sub Reporte_Vista()

On Error GoTo ErrorImpresion
Dim oo As Object

    If OptFecha Or OptRango Then
        MsgBox "Búsqueda no permite impresión, Seleccione Opción x Maquina", vbCritical
        Exit Sub
    End If
    
    strSQL = "cc_muestra_resumen_x_maquina_defectos '" & Trim(TxtCod_MaquinaBus.Text) & "','" & DTPMaquinaDesde.Value & "','" & DTPMaquinaHasta.Value & "'"
    
    sOpcion = "Maq. : " & TxtCod_MaquinaBus & "-" & TxtDes_MaquinaBus & Space(3) & " / Del : " & DTPMaquinaDesde & " al : " & DTPMaquinaHasta
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptAuditoriaTejeduriaVista.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", sOpcion, strSQL, cConnect
    Set oo = Nothing
        
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Sub Reporte_porOT()

On Error GoTo ErrorImpresion
Dim oo As Object

    If Not OptxOT Then
        MsgBox "Búsqueda no permite impresión, Seleccione Opción x OT", vbCritical
        Exit Sub
    End If
    
    strSQL = "cc_muestra_resumen_x_OT '" & Trim(TxtCod_OrdTra.Text) & "'"
    
    sOpcion = "OT : " & TxtCod_OrdTra
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptAuditoriaTejeduriaVista.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", sOpcion, strSQL, cConnect
    Set oo = Nothing
        
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Len(CodOrigen) - PosfinalCod)
End Function


Sub Reporte_porTela()

On Error GoTo ErrorImpresion
Dim oo As Object

    If Not OptTela Then
        MsgBox "Búsqueda no permite impresión, Seleccione Opción x Tela", vbCritical
        Exit Sub
    End If
    
    strSQL = "cc_muestra_resumen_x_Tela '" & Trim(Txtcod_Tela.Text) & "','" & DTPInicio_Tela.Value & "','" & DTPFin_Tela.Value & "'"
    sOpcion = "Tela : " & Txtcod_Tela
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptAuditoriaTejeduriaVista.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", sOpcion, strSQL, cConnect
    Set oo = Nothing
        
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

