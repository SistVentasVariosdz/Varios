VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetalleCrudo_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Crudo"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   10605
      Begin GridEX20.GridEX gexDetalle 
         Height          =   3060
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   5398
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmDetalleCrudo_Ex.frx":0000
         FormatStyle(2)  =   "frmDetalleCrudo_Ex.frx":0138
         FormatStyle(3)  =   "frmDetalleCrudo_Ex.frx":01E8
         FormatStyle(4)  =   "frmDetalleCrudo_Ex.frx":029C
         FormatStyle(5)  =   "frmDetalleCrudo_Ex.frx":0374
         FormatStyle(6)  =   "frmDetalleCrudo_Ex.frx":042C
         FormatStyle(7)  =   "frmDetalleCrudo_Ex.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmDetalleCrudo_Ex.frx":052C
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8160
      TabIndex        =   2
      Top             =   3480
      Width           =   2505
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmDetalleCrudo_Ex.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmDetalleCrudo_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente_Tex As String
Public sSer_OrdComp As String
Public sCod_Ordcomp As String
Dim strSQL As String

Public Sub CARGA_GRID()
On Error GoTo err_Carga
    
    strSQL = "exec ti_sm_trae_guias_crudo_orden_compra_item '" & sCod_Cliente_Tex & "','" & sSer_OrdComp & "','" & sCod_Ordcomp & "'"
    Set gexDetalle.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

Exit Sub
err_Carga:
    ErrorHandler Err, "CARGA_GRID"
End Sub

Sub Reporte()
Dim Scliente As String
On Error GoTo ErrorImpresion
Dim oo As Object

    Scliente = Trim(DevuelveCampo("select nom_cliente from tx_cliente where cod_cliente_tex ='" & sCod_Cliente_Tex & "'", cConnect))
    strSQL = "exec ti_sm_trae_guias_crudo_orden_compra_item '" & sCod_Cliente_Tex & "','" & sSer_OrdComp & "','" & sCod_Ordcomp & "'"

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptDetalleCrudo_OC.XLT"
    oo.Visible = True
    oo.Run "reporte", sSer_OrdComp & "-" & sCod_Ordcomp, Scliente, strSQL, cConnect
    Set oo = Nothing
    
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Guia de Remisión " & Err.Description, vbCritical, "Impresion"

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub
