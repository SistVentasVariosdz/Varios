VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetalleDespachos_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Despachos"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   10215
      Begin GridEX20.GridEX gexDetalle 
         Height          =   3660
         Left            =   60
         TabIndex        =   2
         Top             =   180
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   6456
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
         FormatStyle(1)  =   "frmDetalleDespachos_Ex.frx":0000
         FormatStyle(2)  =   "frmDetalleDespachos_Ex.frx":0138
         FormatStyle(3)  =   "frmDetalleDespachos_Ex.frx":01E8
         FormatStyle(4)  =   "frmDetalleDespachos_Ex.frx":029C
         FormatStyle(5)  =   "frmDetalleDespachos_Ex.frx":0374
         FormatStyle(6)  =   "frmDetalleDespachos_Ex.frx":042C
         FormatStyle(7)  =   "frmDetalleDespachos_Ex.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmDetalleDespachos_Ex.frx":052C
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   9195
      TabIndex        =   0
      Top             =   3930
      Width           =   1065
   End
End
Attribute VB_Name = "frmDetalleDespachos_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SCod_Cliente_Tex As String
Public sSer_OrdComp As String
Public sCod_Ordcomp As String
Public sSec_Ordcomp As String
Dim strSQL As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub CARGA_GRID()
On Error GoTo err_Carga
    
    strSQL = "exec ti_sm_muestra_detalle_despachos_por_orden_compra_item '" & SCod_Cliente_Tex & "','" & sSer_OrdComp & "','" & sCod_Ordcomp & "','" & sSec_Ordcomp & "'"
    Set gexDetalle.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

Exit Sub
err_Carga:
    ErrorHandler Err, "CARGA_GRID"
End Sub


