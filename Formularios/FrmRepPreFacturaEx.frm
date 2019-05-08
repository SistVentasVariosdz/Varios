VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrepprefacturaEx 
   Caption         =   "Reporte de  Prefacturado"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Buscar Por Fecha Emisión"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4815
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3000
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   72089601
            CurrentDate     =   41241
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   72089601
            CurrentDate     =   41241
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   2400
            TabIndex        =   6
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   5520
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   8160
         TabIndex        =   7
         Top             =   240
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmRepPreFacturaEx.frx":0000
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
      Height          =   4815
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   8493
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmRepPreFacturaEx.frx":0090
      FormatStyle(2)  =   "FrmRepPreFacturaEx.frx":01C8
      FormatStyle(3)  =   "FrmRepPreFacturaEx.frx":0278
      FormatStyle(4)  =   "FrmRepPreFacturaEx.frx":032C
      FormatStyle(5)  =   "FrmRepPreFacturaEx.frx":0404
      FormatStyle(6)  =   "FrmRepPreFacturaEx.frx":04BC
      FormatStyle(7)  =   "FrmRepPreFacturaEx.frx":059C
      ImageCount      =   0
      PrinterProperties=   "FrmRepPreFacturaEx.frx":05BC
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   5280
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmrepprefacturaEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsx  As ADODB.Recordset, rsx_Combo As ADODB.Recordset
Dim strsql_x, strsql_x_Combo As String
Private Sub Btn_Buscar_Click()
Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSQL = "Exec Ti_Muestra_Pre_Facturado '" & DTPicker1.Value & "','" & DTPicker2.Value & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

Configura_Grid

GridEX1.ContinuousScroll = True

End Sub
Sub Configura_Grid()

  GridEX1.Columns("FechaTenido").Width = 1000
  GridEX1.Columns("FechaTenido").Caption = "Fecha Teñido"
  
  GridEX1.Columns("Partida").Width = 2000
  GridEX1.Columns("Partida").Caption = "Partida"
  
  GridEX1.Columns("OrdenPedido").Width = 2000
  GridEX1.Columns("OrdenPedido").Caption = "Orden Pedido"
  
  GridEX1.Columns("ClaseOP").Width = 3000
  GridEX1.Columns("ClaseOP").Caption = "Clase"
  

  GridEX1.Columns("cliente").Width = 2000
  GridEX1.Columns("cliente").Caption = "Cliente"
  
  GridEX1.Columns("Tela").Width = 6000
  GridEX1.Columns("Tela").Caption = "Tela"
  
  GridEX1.Columns("Color").Width = 1000
  GridEX1.Columns("Color").Caption = "Color"
  
  GridEX1.Columns("Kilos").Width = 1000
  GridEX1.Columns("Kilos").Caption = "Teñido"
  
  
  GridEX1.Columns("valorizado").Width = 1000
  GridEX1.Columns("valorizado").Caption = "Total"
  
  
End Sub
Private Sub Form_Load()

DTPicker2 = Date
DTPicker1 = DTPicker2 - 30

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
            Case "Imprimir"
                    Call Reporte
            Case "Salir"
                    Unload Me
    End Select
End Sub
Sub Reporte()

    Dim oo As Object
    
    
    If GridEX1.RowCount = 0 Then Exit Sub
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\Rpt_TenidoPreFacturado.xlt"
    oo.Visible = True
    oo.DisplayAlerts = False
         
    
       
    oo.run "REPORTE", DTPicker1.Value, DTPicker2.Value, GridEX1.ADORecordset
    Screen.MousePointer = vbNormal
    'oo.Visible = True
    Set oo = Nothing
    Exit Sub
    
    
End Sub



