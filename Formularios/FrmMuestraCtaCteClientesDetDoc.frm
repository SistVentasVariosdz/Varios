VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraCtaCteClientesDetDoc 
   Caption         =   "DOCUMENTOS CON SALDOS"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1005
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11668
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
      Column(1)       =   "FrmMuestraCtaCteClientesDetDoc.frx":0000
      Column(2)       =   "FrmMuestraCtaCteClientesDetDoc.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmMuestraCtaCteClientesDetDoc.frx":016C
      FormatStyle(2)  =   "FrmMuestraCtaCteClientesDetDoc.frx":02A4
      FormatStyle(3)  =   "FrmMuestraCtaCteClientesDetDoc.frx":0354
      FormatStyle(4)  =   "FrmMuestraCtaCteClientesDetDoc.frx":0408
      FormatStyle(5)  =   "FrmMuestraCtaCteClientesDetDoc.frx":04E0
      FormatStyle(6)  =   "FrmMuestraCtaCteClientesDetDoc.frx":0598
      ImageCount      =   0
      PrinterProperties=   "FrmMuestraCtaCteClientesDetDoc.frx":0678
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&IMPRIMIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1005
   End
End
Attribute VB_Name = "FrmMuestraCtaCteClientesDetDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vanio As String
Public vperiodo As String
Public vcod_anexo As String

Private Sub cmdImprimir_Click()
Call ReportedETALLE
End Sub
Private Sub ReportedETALLE()
Dim oo As Object
Dim Ruta As String, iResp As Integer
Dim sPendCancelTodas  As String


On Error GoTo errReporte

sPendCancelTodas = "P"

'strSQL = "CN_CONSULTA_DUCUMENTOS_VENCIDOS_DETALLE   '$','$','$','$','$'"
'strSQL = VBsprintf(strSQL, vanio, vperiodo, "VN", "C", vcod_anexo)
'Set XRS = CargarRecordSetDesconectado(strSQL, cConnect)

Ruta = vRuta & "\RptSaldosCuentasCorrientesClientesDETALLE.xlt"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = False
oo.displayalerts = False
oo.Run "Reporte", GridEX1.ADORecordset, vanio, vperiodo
oo.Visible = True

Set oo = Nothing

Exit Sub
errReporte:
    MsgBox Err.Description, vbCritical, "Print Voucher Finanzas"
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call mostrar
End Sub
Sub mostrar()
    Dim strSQL As String
    Dim sCodCentroCosto As String
    
    On Error GoTo Fin
    'sCodCentroCosto = dcCentroCostos.BoundText
   
    strSQL = "EXEC CN_CONSULTA_DUCUMENTOS_VENCIDOS_DETALLE '" & vanio & _
                                             "','" & vperiodo & _
                                             "','" & vperiodo & _
                                             "','" & vperiodo & _
                                             "','" & vcod_anexo & "'"
    cadena = strSQL
    
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    Dim C As Integer
        
    'GridEX1.FrozenColumns = 6
    With GridEX1
        
        .Columns("anexo").Visible = False
        .Columns("CLIENTE").Visible = False
        .Columns("num_ruc").Width = 800
        .Columns("num_ruc").Caption = "RUC"
        .Columns("num_ruc").Visible = False
        .Columns("des_anexo").Width = 2000
        .Columns("des_anexo").Caption = "ANEXO"
        .Columns("des_anexo").Visible = False
        
        .Columns("FEC_VENDOC").Width = 1000
        .Columns("FEC_VENDOC").Caption = "VENCIMIENTO"
        .Columns("COD_TIPDOC").Width = 500
        .Columns("COD_TIPDOC").Caption = "TIPO"
        
        .Columns("SER_DOCUM").Width = 500
        .Columns("SER_DOCUM").Caption = "SERIE"
        .Columns("NUM_DOCUM").Width = 1000
        .Columns("NUM_DOCUM").Caption = "NUMERO"
        
        .Columns("COD_MONEDA").Visible = False
        .Columns("SALDO_FINAL").Width = 1000
        .Columns("SALDO_FINAL").Caption = "SALDO SOL"
        .Columns("DOL_SALDO_FINAL").Width = 1000
        .Columns("DOL_SALDO_FINAL").Caption = "SALDO DOL"
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignCenter
        Next C
        
'        With .Columns("COD_ConcepAsist")
'            .Caption = Empty
'            .Width = 600
'            .TextAlignment = jgexAlignLeft
'        End With
'        With .Columns("DESCRIPCION")
'            .Caption = "CONCEPTO DE CONTROL DE ASISTENCIA"
'            .Width = 4000
'            .TextAlignment = jgexAlignLeft
'        End With
'        With .Columns("Cantidad")
'            .Caption = "HORAS"
'            .Width = 1000
'            .TextAlignment = jgexAlignRight
'        End With
        
        
       Dim oGroup01 As GridEX20.JSGroup
'        Dim oGroup02 As GridEX20.JSGroup
'
       Set oGroup01 = .Groups.Add(.Columns("CLIENTE").Index, jgexSortAscending)
'        Set oGroup02 = .Groups.Add(.Columns("TRABAJADOR").Index, jgexSortAscending)
'
        .BackColorRowGroup = &H8000000F
        
'        If CBool(chkExpandir.Value) = True Then
'            .DefaultGroupMode = jgexDGMExpanded
'        Else
'            .DefaultGroupMode = jgexDGMCollapsed
'        End If
'        .ForeColorRowGroup = vbBlue
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
        
        Dim colSOL As JSColumn
        Dim colDOL As JSColumn

        .GroupFooterStyle = jgexTotalsGroupFooter
                  
        Set colSOL = .Columns("SALDO_FINAL")
        With colSOL
            .AggregateFunction = jgexSum
            .TotalRowPrefix = "TOTAL   "
        
        End With
        
        Set colDOL = .Columns("DOL_SALDO_FINAL")
        With colDOL
            .AggregateFunction = jgexSum
            .TotalRowPrefix = ""
        End With
        
        '.SetFocus
    End With
    Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub


