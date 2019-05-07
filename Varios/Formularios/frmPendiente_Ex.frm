VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPendiente_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Servicio Pendientes"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   10605
      Begin GridEX20.GridEX gexDetalle 
         Height          =   3300
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   5821
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
         FormatStyle(1)  =   "frmPendiente_Ex.frx":0000
         FormatStyle(2)  =   "frmPendiente_Ex.frx":0138
         FormatStyle(3)  =   "frmPendiente_Ex.frx":01E8
         FormatStyle(4)  =   "frmPendiente_Ex.frx":029C
         FormatStyle(5)  =   "frmPendiente_Ex.frx":0374
         FormatStyle(6)  =   "frmPendiente_Ex.frx":042C
         FormatStyle(7)  =   "frmPendiente_Ex.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmPendiente_Ex.frx":052C
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8160
      TabIndex        =   2
      Top             =   3720
      Width           =   2505
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmPendiente_Ex.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmPendiente_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente_Tex As String
Public sSer_OrdComp As String
Public sCod_OrdComp As String
Dim StrSQL As String

Public Sub CARGA_GRID()
On Error GoTo err_Carga
    
    StrSQL = "exec ti_sm_trae_orden_servicio_exp"
    Set gexDetalle.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)

Exit Sub
err_Carga:
    ErrorHandler Err, "CARGA_GRID"
End Sub

Sub Reporte()
Dim sCliente As String
On Error GoTo ErrorImpresion
Dim oo As Object

    StrSQL = "exec ti_sm_trae_orden_servicio_exp"

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptDetallePendiente_OC.XLT"
    oo.Visible = True
    oo.Run "reporte", StrSQL, cConnect
    Set oo = Nothing

    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Ordenes de Servicio Pendiente " & Err.Description, vbCritical, "Impresion"

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub
