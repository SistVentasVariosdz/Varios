VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetallePartidas_Ex 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Partidas"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   9720
      TabIndex        =   2
      Top             =   4620
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   4545
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   10965
      Begin GridEX20.GridEX gexDetalle 
         Height          =   4260
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   10830
         _ExtentX        =   19103
         _ExtentY        =   7514
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
         FormatStyle(1)  =   "frmDetallePartidas_Ex.frx":0000
         FormatStyle(2)  =   "frmDetallePartidas_Ex.frx":0138
         FormatStyle(3)  =   "frmDetallePartidas_Ex.frx":01E8
         FormatStyle(4)  =   "frmDetallePartidas_Ex.frx":029C
         FormatStyle(5)  =   "frmDetallePartidas_Ex.frx":0374
         FormatStyle(6)  =   "frmDetallePartidas_Ex.frx":042C
         FormatStyle(7)  =   "frmDetallePartidas_Ex.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmDetallePartidas_Ex.frx":052C
      End
   End
End
Attribute VB_Name = "frmDetallePartidas_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente_Tex As String
Public sSer_OrdComp As String
Public sCod_OrdComp As String
Public sSec_OrdComp As String
Dim StrSql As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub CARGA_GRID()
On Error GoTo err_Carga
    
    StrSql = "exec ti_sm_trae_partida_orden_compra_item '" & sCod_Cliente_Tex & "','" & sSer_OrdComp & "','" & sCod_OrdComp & "','" & sSec_OrdComp & "'"
    Set gexDetalle.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)
    
    Dim fmtCon As JSFmtCondition
 
    Set fmtCon = gexDetalle.FmtConditions.Add(gexDetalle.Columns("Partida").Index, jgexEqual, "TOTAL")
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.ForeColor = &HC00000
Exit Sub
err_Carga:
    ErrorHandler Err, "CARGA_GRID"
End Sub


