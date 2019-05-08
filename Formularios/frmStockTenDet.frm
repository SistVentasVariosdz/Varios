VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmStockTenDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Movimientos Tela Acabada"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   540
      Left            =   9600
      TabIndex        =   1
      Top             =   3840
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexTenDet 
      Height          =   3645
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   6429
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      UseEvenOddColor =   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      RowHeaders      =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmStockTenDet.frx":0000
      FormatStyle(2)  =   "frmStockTenDet.frx":0138
      FormatStyle(3)  =   "frmStockTenDet.frx":01E8
      FormatStyle(4)  =   "frmStockTenDet.frx":029C
      FormatStyle(5)  =   "frmStockTenDet.frx":0374
      FormatStyle(6)  =   "frmStockTenDet.frx":042C
      FormatStyle(7)  =   "frmStockTenDet.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmStockTenDet.frx":052C
   End
End
Attribute VB_Name = "frmStockTenDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String, sCod_Cliente_Tex As String, sCod_Tela As String, _
       sCod_Comb As String, sCod_Color As String, sCod_Talla As String, _
       sCod_Calidad As String, sCod_OrdTra As String
Dim strSQL As String, sTit As String, sErr As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Unload Me
End Sub

Public Sub MostrarDetalleTen()
On Error GoTo ErrDetCru
    
    Me.Caption = "Detalle de Movimiento : " & sCod_OrdTra & "-" & sCod_Tela
    
    Screen.MousePointer = 11
    sTit = "Mostrar Detalle de Movimientos Tela Cruda"

'strSQL = "EXEC TI_SM_KARDEX_MOVIMIENTOS_TELA_TENIDA_CLIENTES '" & _
'             sCod_Almacen & "', '" & sCod_OrdTra & "', '" & sCod_Tela & _
'             "', '" & sCod_Comb & "', '" & sCod_Color & "', '" & _
'             sCod_Talla & "', '" & sCod_Calidad & "'"
             
strSQL = "EXEC SM_MUESTRA_STOCK_DETALLE_ROLLO '" & _
             sCod_Almacen & "', '" & sCod_OrdTra & "', '" & sCod_Tela & _
             "', '" & sCod_Comb & "', '" & sCod_Color & "', '" & _
             sCod_Talla & "', '" & sCod_Calidad & "'"
             
    Set gexTenDet.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
'    gexTenDet.Columns("fec_movstk").Width = 960
'    gexTenDet.Columns("Tipo_Mov").Width = 2400
'    gexTenDet.Columns("Tela").Width = 1920
'    gexTenDet.Columns("Comb").Width = 240
'    gexTenDet.Columns("Color").Width = 1725
'    gexTenDet.Columns("Talla").Width = 465
'    gexTenDet.Columns("Calidad").Width = 375
'    gexTenDet.Columns("kgs_movimiento").Width = 1005
'    gexTenDet.Columns("uni_movimiento").Width = 915
'    gexTenDet.Columns("numero_rollos").Width = 885
'    gexTenDet.Columns("Proveedor").Width = 915
'    gexTenDet.Columns("Guia").Width = 1110
'    gexTenDet.Columns("Parte_Salida").Width = 1035
'    gexTenDet.Columns("Observaciones").Width = 1200
'    gexTenDet.Columns("num_movstk").Width = 1035
'    gexTenDet.Columns("cod_maquina_tinto").Width = 705
'
'    gexTenDet.Columns("fec_movstk").Caption = "Fecha"
'    gexTenDet.Columns("Tipo_Mov").Caption = "Tipo Mov."
'    gexTenDet.Columns("Tela").Caption = "Tela"
'    gexTenDet.Columns("Comb").Caption = "Comb"
'    gexTenDet.Columns("Color").Caption = "Color"
'    gexTenDet.Columns("Talla").Caption = "Talla"
'    gexTenDet.Columns("Calidad").Caption = "Cal"
'    gexTenDet.Columns("kgs_movimiento").Caption = "Kgs.Mov."
'    gexTenDet.Columns("uni_movimiento").Caption = "Und.Mov."
'    gexTenDet.Columns("numero_rollos").Caption = "Nro.Rollos"
'    gexTenDet.Columns("Proveedor").Caption = "Proveedor"
'    gexTenDet.Columns("Guia").Caption = "Guia"
'    gexTenDet.Columns("Parte_Salida").Caption = "Part.Sal"
'    gexTenDet.Columns("Observaciones").Caption = "Observaciones"
'    gexTenDet.Columns("num_movstk").Caption = "Num.Mov"
'    gexTenDet.Columns("cod_maquina_tinto").Caption = "Maquina"
'
    Screen.MousePointer = 0
Exit Sub
ErrDetCru:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
