VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmConsultaMovTelaTenida 
   Caption         =   "Movimientos Tela Acabada"
   ClientHeight    =   6480
   ClientLeft      =   3480
   ClientTop       =   2460
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   13305
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   700
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13215
      Begin VB.TextBox txtPartida 
         Height          =   285
         Left            =   5760
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3975
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   11880
         TabIndex        =   3
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker DTPFechaHasta 
         Height          =   315
         Left            =   10080
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   78708737
         CurrentDate     =   38169
      End
      Begin MSComCtl2.DTPicker DTPFechaDesde 
         Height          =   315
         Left            =   7680
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   78708737
         CurrentDate     =   38169
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Partida:"
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hasta"
         Height          =   195
         Left            =   9480
         TabIndex        =   9
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desde"
         Height          =   195
         Left            =   7080
         TabIndex        =   8
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Almacén"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   13215
      Begin GridEX20.GridEX gexLista 
         Height          =   4845
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   13020
         _ExtentX        =   22966
         _ExtentY        =   8546
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "FrmConsultaMovTelaTenida.frx":0000
         FormatStyle(2)  =   "FrmConsultaMovTelaTenida.frx":0138
         FormatStyle(3)  =   "FrmConsultaMovTelaTenida.frx":01E8
         FormatStyle(4)  =   "FrmConsultaMovTelaTenida.frx":029C
         FormatStyle(5)  =   "FrmConsultaMovTelaTenida.frx":0374
         FormatStyle(6)  =   "FrmConsultaMovTelaTenida.frx":042C
         FormatStyle(7)  =   "FrmConsultaMovTelaTenida.frx":050C
         ImageCount      =   0
         PrinterProperties=   "FrmConsultaMovTelaTenida.frx":052C
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   10560
      TabIndex        =   10
      Top             =   5880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   900
      Custom          =   $"FrmConsultaMovTelaTenida.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1250
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   6000
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmConsultaMovTelaTenida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rstAux As ADODB.Recordset
Dim scod_almacen, sDes_Almacen As String
Dim tipo As String

Public accion As String

Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
FillAlmacen
Me.DTPFechaDesde.Value = Date
Me.DTPFechaHasta.Value = Date
tipo = "0"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
BUSCAR ""
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR "
    BUSCAR ""
    Reporte
Case "INGRESOS"
    BUSCAR "1"
    tipo = "1"
    Reporte
Case "IMPRIMIROTROS"
    BUSCAR "2"
    tipo = "2"
    Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Sub BUSCAR(accion As String)
On Error GoTo ErrPesadas

    If cboAlmacen.ListIndex = -1 Then
        MsgBox "Se debe elegir un Almacen", vbExclamation + vbOKOnly, "Almacenes"
        Exit Sub
    End If
    scod_almacen = Left(cboAlmacen, 2)
    sDes_Almacen = Trim(Mid(cboAlmacen, 4))
    
    Screen.MousePointer = 11
    strSQL = "EXEC ti_muestra_movimientos_tela_tenida_intervalo '" & scod_almacen & "', '" & _
             DTPFechaDesde.Value & "', '" & DTPFechaHasta.Value & "','" & accion & "','" & txtPartida.Text & "'"
             
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    gexLista.Columns("Num_MovStk").Width = 1110
    gexLista.Columns("Num_MovStk").Caption = "Num MovStk"
    gexLista.Columns("fec_movstk").Width = 960
    gexLista.Columns("fec_movstk").Caption = "Fec Movstk"
    gexLista.Columns("Tipo_Movimiento").Width = 1995
    gexLista.Columns("Tipo_Movimiento").Caption = "Tipo Movimiento"
    gexLista.Columns("Nom_Cliente").Width = 1035
    gexLista.Columns("Nom_Cliente").Caption = "Nom Cliente"
    gexLista.Columns("Proveedor").Width = 2475
    gexLista.Columns("Proveedor").Caption = "Proveedor"
    gexLista.Columns("Guia").Width = 1110
    gexLista.Columns("Guia").Caption = "Guia"
    gexLista.Columns("Partida").Width = 630
    gexLista.Columns("Partida").Caption = "Partida"
    gexLista.Columns("Codigo").Width = 885
    gexLista.Columns("Codigo").Caption = "Codigo"
    gexLista.Columns("Nombre_Tela").Width = 2100
    gexLista.Columns("Nombre_Tela").Caption = "Nombre Tela"
    gexLista.Columns("Comb").Width = 540
    gexLista.Columns("Comb").Caption = "Comb"
    gexLista.Columns("Des_Comb").Width = 2145
    gexLista.Columns("Des_Comb").Caption = "Des Comb"
    gexLista.Columns("Color").Width = 660
    gexLista.Columns("Color").Caption = "Color"
    gexLista.Columns("Nombre_Color").Width = 1770
    gexLista.Columns("Nombre_Color").Caption = "Nombre Color"
    gexLista.Columns("Talla").Width = 690
    gexLista.Columns("Talla").Caption = "Talla"
    gexLista.Columns("Cal").Width = 360
    gexLista.Columns("Cal").Caption = "Cal"
    gexLista.Columns("Kgs").Width = 480
    gexLista.Columns("Kgs").Caption = "Kgs"
    gexLista.Columns("Rollos").Width = 570
    gexLista.Columns("Rollos").Caption = "Rollos"
    gexLista.Columns("Orden_Compra").Width = 1200
    gexLista.Columns("Orden_Compra").Caption = "Orden Compra"
    gexLista.Columns("kgs_segun_guia").Width = 1305
    gexLista.Columns("kgs_segun_guia").Caption = "Kgs Guia"
    gexLista.Columns("nro_rollos_segun_guia").Width = 1725
    gexLista.Columns("nro_rollos_segun_guia").Caption = "Nro Rollos-Guia"
    gexLista.Columns("Observaciones").Width = 2325
    gexLista.Columns("Observaciones").Caption = "Observaciones"
    
    gexLista.FrozenColumns = 3
        
    Screen.MousePointer = 0
Exit Sub
ErrPesadas:
    ErrorHandler err, "Buscar"
    Screen.MousePointer = 0
End Sub


Sub Reporte()
On Error GoTo ErrorImpresion
    Dim oo As Object
    
    
    If gexLista.RowCount = 0 Then Exit Sub
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\TI_MovTelaTenida.xlt"
    oo.Visible = True
    oo.DisplayAlerts = False
    
       
    oo.Run "REPORTE", scod_almacen, sDes_Almacen, Me.DTPFechaDesde.Value, Me.DTPFechaHasta.Value, gexLista.ADORecordset, tipo
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Consulta de Mov. Tela Teñida " & err.Description, vbCritical, "Impresion"
End Sub

Private Sub FillAlmacen()

On Error GoTo Fin
Dim sTit As String
    
    sTit = "Cargar Almacenes"
    
    strSQL = "SELECT Cod_Almacen, Nom_Almacen FROM TX_ALMACEN " & _
             "WHERE  Tip_Item = 'T' " & _
             "AND    Tip_Presentacion = 'T' "
    
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cboAlmacen.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cboAlmacen.AddItem !Cod_almacen & " " & !nom_almacen
            .MoveNext
        Loop
        .Close
    End With
    If cboAlmacen.ListCount > 0 Then cboAlmacen.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
    
End Sub


