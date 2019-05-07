VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraCtaCteClientesRangos 
   Caption         =   "SALDOS CTA CTE CLIENTES"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimirDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "IMPRIMIR DETALLE"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1005
   End
   Begin VB.CommandButton cmdVerDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&VER DETALLE"
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
      TabIndex        =   9
      Top             =   7200
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   10860
      Begin VB.TextBox txtAno 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1425
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txtPeriodo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.TextBox txtCod_anexo 
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txtDes_Anexo 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   5565
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&BUSCAR"
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
         Left            =   9720
         Picture         =   "FrmMuestraCtaCteClientesRangos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "ANEXO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   8
         Tag             =   "Document Type"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "AÑO/PERIODO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Tag             =   "Document Type"
         Top             =   270
         Width           =   1230
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11033
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
      Column(1)       =   "FrmMuestraCtaCteClientesRangos.frx":0102
      Column(2)       =   "FrmMuestraCtaCteClientesRangos.frx":01CA
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmMuestraCtaCteClientesRangos.frx":026E
      FormatStyle(2)  =   "FrmMuestraCtaCteClientesRangos.frx":03A6
      FormatStyle(3)  =   "FrmMuestraCtaCteClientesRangos.frx":0456
      FormatStyle(4)  =   "FrmMuestraCtaCteClientesRangos.frx":050A
      FormatStyle(5)  =   "FrmMuestraCtaCteClientesRangos.frx":05E2
      FormatStyle(6)  =   "FrmMuestraCtaCteClientesRangos.frx":069A
      ImageCount      =   0
      PrinterProperties=   "FrmMuestraCtaCteClientesRangos.frx":077A
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   1005
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6240
      Top             =   7320
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMuestraCtaCteClientesRangos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSQL   As String
Public codigo As String
Public Descripcion As String
Private Sub cmdBuscar_Click()

    GridEX1.ClearFields
    Call mostrar
    
End Sub

Sub mostrar()
    Dim strSQL As String
    Dim sCodCentroCosto As String
    
    On Error GoTo Fin
    'sCodCentroCosto = dcCentroCostos.BoundText
   
    strSQL = "EXEC CN_Consulta_ducumentos_vencidos '" & txtAno.Text & _
                                             "','" & txtPeriodo.Text & _
                                             "','" & txtPeriodo.Text & _
                                             "','" & txtPeriodo.Text & _
                                             "','" & txtCod_anexo.Text & "'"
    cadena = strSQL
    
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    Dim C As Integer
        
    'GridEX1.FrozenColumns = 6
    With GridEX1
        
        .Columns("anexo").Visible = False
        .Columns("num_ruc").Width = 800
        .Columns("des_anexo").Width = 2000
        
       .Columns("sal_sol_h00").Width = 795
.Columns("sal_dol_h00").Width = 795

.Columns("sal_sol_h30").Width = 795
.Columns("sal_dol_h30").Width = 795

.Columns("sal_sol_h60").Width = 795
.Columns("sal_dol_h60").Width = 795

.Columns("sal_sol_h90").Width = 795
.Columns("sal_dol_h90").Width = 795

.Columns("sal_sol_h120").Width = 795
.Columns("sal_dol_h120").Width = 795

.Columns("sal_sol_h150").Width = 795
.Columns("sal_dol_h150").Width = 795

.Columns("sal_sol_h180").Width = 795
.Columns("sal_dol_h180").Width = 795

.Columns("sal_sol_h210").Width = 795
.Columns("sal_dol_h210").Width = 795

.Columns("sal_sol_h240").Width = 795
.Columns("sal_dol_h240").Width = 795

.Columns("sal_sol_h270").Width = 795
.Columns("sal_dol_h270").Width = 795

.Columns("sal_sol_h300").Width = 795
.Columns("sal_dol_h300").Width = 795

.Columns("sal_sol_h330").Width = 795
.Columns("sal_dol_h330").Width = 795

.Columns("sal_sol_h360").Width = 795
.Columns("sal_dol_h360").Width = 795
        
.Columns("SAL_SOL_H00").Caption = "SS_00"
.Columns("SAL_DOL_H00").Caption = "S$_00"

.Columns("SAL_SOL_H30").Caption = "SS_30"
.Columns("SAL_DOL_H30").Caption = "S$_30"

.Columns("SAL_SOL_H60").Caption = "SS_60"
.Columns("SAL_DOL_H60").Caption = "S$_60"

.Columns("SAL_SOL_H90").Caption = "SS_90"
.Columns("SAL_DOL_H90").Caption = "S$_90"

.Columns("SAL_SOL_H120").Caption = "SS_120"
.Columns("SAL_DOL_H120").Caption = "S$_120"

.Columns("SAL_SOL_H150").Caption = "SS_150"
.Columns("SAL_DOL_H150").Caption = "S$_150"

.Columns("SAL_SOL_H180").Caption = "SS_180"
.Columns("SAL_DOL_H180").Caption = "S$_180"

.Columns("SAL_SOL_H210").Caption = "SS_210"
.Columns("SAL_DOL_H210").Caption = "S$_210"

.Columns("SAL_SOL_H240").Caption = "SS_240"
.Columns("SAL_DOL_H240").Caption = "S$_240"

.Columns("SAL_SOL_H270").Caption = "SS_270"
.Columns("SAL_DOL_H270").Caption = "S$_270"

.Columns("SAL_SOL_H300").Caption = "SS_300"
.Columns("SAL_DOL_H300").Caption = "S$_300"

.Columns("SAL_SOL_H330").Caption = "SS_330"
.Columns("SAL_DOL_H330").Caption = "S$_330"

.Columns("SAL_SOL_H360").Caption = "SS_360"
.Columns("SAL_DOL_H360").Caption = "S$_360"
        
        
.Columns("total_sol").Width = 1000
.Columns("total_sol").Caption = "TOTAL S"
.Columns("Total_dol").Width = 1000
.Columns("Total_dol").Caption = "TOTAL $"
        
        
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
        
        
'        Dim oGroup01 As GridEX20.JSGroup
'        Dim oGroup02 As GridEX20.JSGroup
'
'        Set oGroup01 = .Groups.Add(.Columns("CENTRO_COSTO").Index, jgexSortAscending)
'        Set oGroup02 = .Groups.Add(.Columns("TRABAJADOR").Index, jgexSortAscending)
'
'        .BackColorRowGroup = &H8000000F
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
        
'        Dim colHORAS As JSColumn
'
'        .GroupFooterStyle = jgexTotalsGroupFooter
'
'        Set colHORAS = .Columns("CANTIDAD")
'        With colHORAS
'            .AggregateFunction = jgexSum
'            .TotalRowPrefix = ""
'        End With
        
        .SetFocus
    End With
    Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub cmdImprimir_Click()
Call Reporte
End Sub
Private Sub Reporte()
Dim oo As Object
Dim Ruta As String, iResp As Integer
Dim sPendCancelTodas  As String
On Error GoTo errReporte

sPendCancelTodas = "P"

'strSQL = "CN_Consulta_ducumentos_vencidos   '$','$','$','$','$','$'"
'strSQL = VBsprintf(strSQL, txtAno.Text, txtPeriodo.Text, "VN", "C", txtCod_Anxo.Text)

Ruta = vRuta & "\RptSaldosCuentasCorrientesClientes.xlt"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = False
oo.displayalerts = False
oo.Run "Reporte", GridEX1.ADORecordset, Trim(txtAno.Text), Trim(txtPeriodo.Text)
oo.Visible = True

Set oo = Nothing

Exit Sub
errReporte:
    MsgBox Err.Description, vbCritical, "Print Voucher Finanzas"
End Sub
Private Sub ReportedETALLE()
Dim oo As Object
Dim Ruta As String, iResp As Integer
Dim sPendCancelTodas  As String
Dim XRS As New ADODB.Recordset

On Error GoTo errReporte

sPendCancelTodas = "P"

strSQL = "CN_CONSULTA_DUCUMENTOS_VENCIDOS_DETALLE   '$','$','$','$','$'"
strSQL = VBsprintf(strSQL, txtAno.Text, txtPeriodo.Text, "VN", "C", txtCod_anexo.Text)

Set XRS = CargarRecordSetDesconectado(strSQL, cConnect)
Ruta = vRuta & "\RptSaldosCuentasCorrientesClientesDETALLE.xlt"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = False
oo.displayalerts = False
oo.Run "Reporte", XRS, Trim(txtAno.Text), Trim(txtPeriodo.Text)
oo.Visible = True

Set oo = Nothing

Exit Sub
errReporte:
    MsgBox Err.Description, vbCritical, "Print Voucher Finanzas"
End Sub


Private Sub cmdImprimirDetalle_Click()
Call ReportedETALLE
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVerDetalle_Click()

If Right(Trim(GridEX1.Value(GridEX1.Columns("ANEXO").Index)), 4) = "" Then Exit Sub
If GridEX1.RowCount = 0 Then Exit Sub
FrmMuestraCtaCteClientesDetDoc.vanio = txtAno.Text
FrmMuestraCtaCteClientesDetDoc.vperiodo = txtPeriodo
FrmMuestraCtaCteClientesDetDoc.vcod_anexo = Right(Trim(GridEX1.Value(GridEX1.Columns("ANEXO").Index)), 4)
'FrmMuestraCtaCteClientesDetDoc.mostrar
Load FrmMuestraCtaCteClientesDetDoc
FrmMuestraCtaCteClientesDetDoc.Show 1

End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPeriodo.SetFocus
End If

End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call Busca_Opcion("Cod_Anxo", "Des_Anexo", "CN_AnexosContables WHERE Cod_TipAnEX ='C' AND ", txtCod_anexo, txtDes_Anexo, 2, Me)
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPeriodo.Text = Right("00" + txtPeriodo.Text, 2)
    txtCod_anexo.SetFocus
End If

End Sub
Private Sub txtCod_anexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Anxo", "Des_Anexo", "CN_AnexosContables WHERE Cod_TipAnEX ='C' AND ", txtCod_anexo, txtDes_Anexo, 1, Me)
    txtDes_Anexo.SetFocus
End If
End Sub


Public Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String

    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    
    
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
   
    
    End Select
    txtCod = ""
    txtDes = ""
    
    With frmBusqGeneral
        Set .oParent = frmME
        .sQuery = strSQL
        .Cargar_Datos
        
        frmME.codigo = ""
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.codigo = ".."
        End If
        
        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod = frmME.codigo 'Trim(rstAux!Cod)
            txtDes = frmME.Descripcion  'Trim(rstAux!Descripcion)
            
            If txtCod = ".." Or frmME.Descripcion = "" Then
                txtCod = Trim(rstAux!Cod)
                txtDes = Trim(rstAux!Descripcion)
            End If
            
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
            
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub



